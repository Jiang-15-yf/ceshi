"""core.excel_pricecard - Excel/CSV 宽表 → 特殊价格牌（竖排分组格式）

字段约定（与原 HTML/JS 一致）：
- 输入宽表列：款号｜颜色｜背面｜尺码段｜XS…XXL｜合计｜涉及PO
- 输出竖排表列：[颜色, 尺码, 订单数, 生产数, 尺码段, 背面, PO号, 款号]
- S 码行绿色（#CCFFCC），合计行黄色（#FFFF00），空白吊牌汇总为单行
- 上方 A1 起保留原始宽表供对照；右侧 L10 起嵌入带尺寸标注的标签参考图
"""
from __future__ import annotations

import csv
import math
import re
from dataclasses import dataclass
from io import BytesIO
from typing import Dict, List, Any, Optional, Tuple

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 合法尺码集合（原 JS 同款）
SIZE_SET = {"XXS", "XS", "S", "M", "L", "XL", "XXL", "2XL", "3XL", "XXXL", "4XL", "5XL"}

# Excel 单元格样式
THIN = Side(border_style="thin", color="808080")
BORDER = Border(top=THIN, left=THIN, bottom=THIN, right=THIN)
FONT_NAME = "宋体"
FONT_SIZE = 10

# 颜色（与原 JS 一致：去前缀的 6 位十六进制）
FILL_TOTAL = "FFFFFF00"  # 黄色
FILL_S = "FFCCFFCC"      # 浅绿


# ================================================================
# 读取与解析
# ================================================================

def _norm(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _to_num(v: Any) -> float:
    if isinstance(v, (int, float)):
        return v
    if v is None or v == "":
        return 0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0


def _read_aoa(path: str) -> List[List[Any]]:
    """读取 Excel 或 CSV 为二维数组（AOA）

    按后缀自动选引擎：
    - .csv → Python 内置 csv
    - .xls → xlrd 1.2.0（最后支持 .xls 的版本；需 `pip install xlrd==1.2.0`）
    - .xlsx → openpyxl
    """
    p = path.lower()
    if p.endswith(".csv"):
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            return [row for row in reader]
    if p.endswith(".xls"):
        import xlrd  # 仅 .xls 路径下导入，避免污染 .xlsx 用户环境
        wb = xlrd.open_workbook(path, formatting_info=False)
        ws = wb.sheet_by_index(0)
        aoa: List[List[Any]] = []
        for r in range(ws.nrows):
            row: List[Any] = []
            for c in range(ws.ncols):
                cell = ws.cell(r, c)
                # xlrd 把空单元格当 ''，日期/数字也照常返回；保持与 openpyxl 行为一致
                row.append(cell.value)
            aoa.append(row)
        return aoa
    # 默认 .xlsx
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    return [[c.value for c in row] for row in ws.iter_rows()]


def parse_table(path: str) -> Dict[str, Any]:
    """解析宽表 → {header, rows, size_headers}

    rows 每条记录字段：
        style, color, back, seg, po, sizes(list[float]), total(float)
    """
    data = _read_aoa(path)
    # 定位表头行
    h_row = -1
    for i in range(min(len(data), 20)):
        cells = [_norm(c) for c in data[i]]
        if "款号" in cells and "颜色" in cells:
            h_row = i
            break
    if h_row < 0:
        raise ValueError("未找到表头行（需同时包含「款号」和「颜色」列）")
    H = [_norm(c) for c in data[h_row]]
    c_style = H.index("款号")
    c_color = H.index("颜色")
    c_back = H.index("背面") if "背面" in H else -1
    c_seg = H.index("尺码段") if "尺码段" in H else -1
    c_po = next((i for i, h in enumerate(H) if "PO" in h), -1)
    if c_style < 0 or c_color < 0:
        raise ValueError("表头缺少「款号/颜色」列")

    lo = (c_seg if c_seg >= 0 else 3) + 1
    hi = c_po if c_po >= 0 else len(H)
    size_cols: List[int] = []
    total_col = -1
    for c in range(lo, hi):
        h = H[c].upper().replace(" ", "")
        if h in SIZE_SET:
            size_cols.append(c)
        elif H[c] != "":
            total_col = c
    if total_col < 0 and hi - 1 > size_cols[-1]:
        total_col = size_cols[-1] + 1

    rows: List[Dict[str, Any]] = []
    for r in range(h_row + 1, len(data)):
        line = data[r] or []
        joined = "|".join(_norm(c) for c in line)
        if not joined.replace("|", "").strip():
            continue
        if "合计" in joined:
            break

        def get(c: int) -> Any:
            return line[c] if 0 <= c < len(line) else None

        rec = {
            "style": _norm(get(c_style)),
            "color": _norm(get(c_color)),
            "back": _norm(get(c_back)),
            "seg": _norm(get(c_seg)),
            "po": _norm(get(c_po)),
            "sizes": [_to_num(get(c)) for c in size_cols],
            "total": _to_num(get(total_col)) if total_col >= 0 else 0,
        }
        if not rec["style"] and not rec["color"] and all(s == 0 for s in rec["sizes"]):
            continue
        rows.append(rec)

    if not rows:
        raise ValueError("未读取到数据行")
    size_headers = [H[c].upper().replace(" ", "") for c in size_cols]
    return {"header": H, "rows": rows, "size_headers": size_headers}


# ================================================================
# 构建竖排数据
# ================================================================

def build_pricecard(parsed: Dict[str, Any], rate_pct: float) -> Dict[str, Any]:
    """按上浮比例构建竖排数据

    返回：
        body       : 主体行（每个尺码一行）
        blank_out  : 空白吊牌汇总
        total_row  : 合计行
    """
    rate = 1 + (rate_pct or 0) / 100
    prod = lambda n: math.ceil((n or 0) * rate - 1e-9)

    rows = parsed["rows"]
    priced = [r for r in rows if "有价格" in r["back"]]
    blanks = [r for r in rows if "空白" in r["back"]]
    use_p = priced if priced else rows

    # 主体行
    body: List[List[Any]] = []
    for r in use_p:
        for i, q in enumerate(r["sizes"]):
            body.append([
                r["color"], parsed["size_headers"][i], q, prod(q),
                r["seg"], r["back"], r["po"], r["style"],
            ])

    # 空白吊牌汇总（按款号聚合）
    blank_groups: Dict[str, Dict[str, Any]] = {}
    for r in blanks:
        g = blank_groups.setdefault(r["style"], {"order": 0, "po": set()})
        g["order"] += sum(r["sizes"])
        if r["po"]:
            for p in re.split(r"[,，、]", r["po"]):
                g["po"].add(p.strip())
    blank_out = []
    for style, g in blank_groups.items():
        blank_out.append([
            "空白吊牌", "", g["order"], prod(g["order"]),
            "", "背面空白", ",".join(p for p in g["po"] if p), style,
        ])

    # 合计
    tot_o = sum(r[2] for r in body) + sum(r[2] for r in blank_out)
    tot_p = sum(r[3] for r in body) + sum(r[3] for r in blank_out)
    total_row = ["合计", "", tot_o, tot_p, "", "", "", ""]

    return {"body": body, "blank_out": blank_out, "total_row": total_row}


# ================================================================
# 工作表模型（导出与预览共用，保证「预览 = 实际输出」）
# ================================================================

@dataclass
class SheetCell:
    value: Any = None
    fill: Optional[str] = None        # ARGB 十六进制，如 "FFFFFF00"
    bold: bool = False
    align: str = "center"             # "center" / "left" / "right"


@dataclass
class SheetModel:
    n_rows: int
    n_cols: int
    cells: Dict[Tuple[int, int], SheetCell]                       # (row, col)，均 1-based
    merges: List[Tuple[int, int, int, int]]                       # (r1, c1, r2, c2) 含端点，1-based
    col_widths: Dict[int, float]                                  # col(1-based) -> 字符宽
    row_heights: Dict[int, float]                                 # row(1-based) -> 点高
    images: List[Tuple[str, bytes, float, float]]                 # (anchor_cell, png_bytes, width_px, height_px)


def build_sheet_model(
    parsed: Dict[str, Any],
    built: Dict[str, Any],
    tag_w_cm: float = 4.5,
    tag_h_cm: float = 9,
    ref_image_bytes: Optional[bytes] = None,
) -> SheetModel:
    """构建与实际导出 100% 一致的工作表模型。

    布局：
    - A1 起：原始宽表（表头 + 数据行）
    - D10 起：竖排特殊价格牌（表头 + 主体 + 空白吊牌 + 合计）
    - L10 起：标签参考图（带尺寸标注）
    """
    from .label_image import make_annotated_from_bytes, make_label_diagram

    cells: Dict[Tuple[int, int], SheetCell] = {}
    merges: List[Tuple[int, int, int, int]] = []

    def put(r: int, c: int, v: Any, fill: Optional[str] = None, bold: bool = False, align: str = "center"):
        # 整值浮点转 int，确保与 openpyxl 写回的整数一致（26.0 → 26）
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        cells[(r, c)] = SheetCell(value=v, fill=fill, bold=bold, align=align)

    # --- 上方：原始宽表（A1 起） ---
    H = parsed["header"]
    for c, h in enumerate(H, 1):
        if h:
            put(1, c, h, bold=True, align="left")
    for i, r in enumerate(parsed["rows"], 2):
        put(i, 1, r["style"], align="left")
        put(i, 2, r["color"], align="left")
        put(i, 3, r["back"], align="left")
        put(i, 4, r["seg"], align="left")
        for j, q in enumerate(r["sizes"]):
            put(i, 5 + j, q)
        put(i, 5 + len(r["sizes"]), r["total"])
        put(i, 6 + len(r["sizes"]), r["po"], align="left")

    # --- 下方：竖排特殊价格牌（D10 起） ---
    start_row = 10
    start_col = 4  # D 列
    heads = ["颜色", "尺码", "订单数", "生产数", "尺码段", "背面", "PO号", "款号"]
    for i, h in enumerate(heads):
        put(start_row, start_col + i, h, bold=True, align="center")

    all_rows = built["body"] + built["blank_out"] + [built["total_row"]]
    last_idx = len(all_rows) - 1
    body_len = len(built["body"])
    for ri, row in enumerate(all_rows):
        R = start_row + 1 + ri
        is_total = ri == last_idx
        is_blank = body_len <= ri < last_idx
        if is_total:
            fill = FILL_TOTAL
        elif not is_blank and str(row[1]).upper() == "S":
            fill = FILL_S
        else:
            fill = None
        for ci, v in enumerate(row):
            put(R, start_col + ci, v, fill=fill, bold=is_total,
                align="left" if ci in (0, 4, 5, 6, 7) else "center")

    # 合并相同值列（颜色/尺码段/背面/PO号/款号）— 与导出完全一致
    merge_cols = [0, 4, 5, 6, 7]
    for ci in merge_cols:
        s = 0
        for i in range(1, len(all_rows) + 1):
            cur = str(all_rows[i][ci]) if i < len(all_rows) else None
            prev = str(all_rows[s][ci])
            if cur != prev:
                if prev != "" and i - s > 1:
                    merges.append((start_row + 1 + s, start_col + ci,
                                   start_row + i, start_col + ci))
                s = i

    # 列宽。未显式设置的列回落到与 openpyxl 一致的默认 13.0（预览/导出同源）。
    col_widths: Dict[int, float] = {c: 13.0 for c in range(1, 100)}
    vcard_widths = [12, 8, 9, 9, 10, 10, 18, 16]
    for i, w in enumerate(vcard_widths):
        col_widths[start_col + i] = w
    col_widths[12] = 24
    col_widths[13] = 24
    # 宽表各列给能看清楚的宽度
    if parsed["rows"]:
        wtable_widths = [14, 10, 14, 12] + [8] * len(parsed["rows"][0]["sizes"]) + [9, 20]
        for i, w in enumerate(wtable_widths):
            col_widths[i + 1] = w

    # 行高
    row_heights: Dict[int, float] = {}
    for r in range(1, 31):
        row_heights[r] = 24 if start_row <= r <= 30 else 15
    # 宽表表头/数据行高
    for r in range(1, 2 + len(parsed["rows"])):
        row_heights.setdefault(r, 20)

    n_rows = max(2 + len(parsed["rows"]), start_row + len(all_rows), 30)
    n_cols = max(13, len(H) + 1)

    # 参考图
    images: List[Tuple[str, bytes, float, float]] = []
    try:
        if ref_image_bytes:
            img_bytes = make_annotated_from_bytes(ref_image_bytes, tag_w_cm, tag_h_cm)
        else:
            img_bytes = make_label_diagram(tag_w_cm, tag_h_cm)
        images.append(("L10", img_bytes, 240, 480))
    except Exception as e:
        print(f"[warn] 生成标签图失败: {e}")

    return SheetModel(
        n_rows=n_rows, n_cols=n_cols, cells=cells, merges=merges,
        col_widths=col_widths, row_heights=row_heights, images=images,
    )


def _write_model(model: SheetModel, output_path: str) -> None:
    """把 SheetModel 实际写出为 xlsx（与预览使用同一模型 → 完全一致）"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.sheet_view.zoomScale = 100

    FONT_NAME = "宋体"
    FONT_SIZE = 10
    for (r, c), cell in model.cells.items():
        xl = ws.cell(row=r, column=c, value=cell.value)
        xl.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=cell.bold)
        xl.alignment = Alignment(
            horizontal=cell.align, vertical="center", wrap_text=True)
        xl.border = BORDER
        if cell.fill:
            xl.fill = PatternFill("solid", fgColor=cell.fill)

    for (r1, c1, r2, c2) in model.merges:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    for c, w in model.col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    for r, h in model.row_heights.items():
        ws.row_dimensions[r].height = h

    for anchor, img_bytes, w, h in model.images:
        img = XLImage(BytesIO(img_bytes))
        img.width = int(w)
        img.height = int(h)
        ws.add_image(img, anchor)

    wb.save(output_path)


def export_xlsx(
    parsed: Dict[str, Any],
    built: Dict[str, Any],
    output_path: str,
    tag_w_cm: float = 4.5,
    tag_h_cm: float = 9,
    ref_image_bytes: Optional[bytes] = None,
) -> None:
    """导出 .xlsx：内部先构建 SheetModel，再用 _write_model 写出。
    预览与导出共用同一 SheetModel，确保「预览所见即所得」。"""
    model = build_sheet_model(
        parsed, built, tag_w_cm=tag_w_cm, tag_h_cm=tag_h_cm,
        ref_image_bytes=ref_image_bytes,
    )
    _write_model(model, output_path)
