"""core.pdf_rearrange - PDF 价格签自动检测与重排

核心算法（与原 HTML/JS 版本对齐）：
1. render_page_pix  : PyMuPDF 把页面渲染成高分辨率 RGB 位图（DETECT_SCALE=3）
2. detect_stickers  : 行/列投影 + 阈值 RGB<235 → 候选矩形 → merge_boxes 合并 → 加 PAD
3. merge_boxes      : 先按 x 中心聚列，再按 y 聚行，最后按 y 全局二次合并
4. rearrange_pdf    : 把检测到的标签按 PDF 源页面区域直接绘制到新页面（保清晰度）

居中与尺寸统一策略（v1.3 起）：
- 全部标签的标称尺寸取「中位数 + 容差吸收抖动」，得到一个统一的 crop_w × crop_h；
- 每枚标签的裁切框 = 以它自己的墨迹中心 (cx, cy) 为圆心、尺寸恒为 crop_w × crop_h；
- 目标槽位尺寸与裁切框相同，缩放恒为 1.0，直接居中摆放。
结果：所有输出标签尺寸 100% 一致，且每枚标签的墨迹中心严格落在输出槽位中心。
裁切框越出源页边界时（标签贴边），只绘制页面内的部分，越界处留白，居中关系不变。

单位约定：
- 检测坐标：DETECT_SCALE=3 下像素（仅检测阶段使用）
- 输出坐标：PDF 点（pt，72pt=1inch），与 PyMuPDF 一致
- 转换：pt = px / DETECT_SCALE
"""
from __future__ import annotations

import math
import statistics
from dataclasses import dataclass
from typing import List, Tuple, Dict, Any

import fitz  # PyMuPDF
from PIL import Image

from .units import mm_to_pt, pt_to_cm

# 检测固定放大倍率（与原 JS 一致，DETECT_SCALE 像素→pt 需除以此值）
DETECT_SCALE = 3
# 检测框外扩像素（≈ 10pt ≈ 3.5mm，避免切边）
PAD_PX = 30
# 墨色阈值（原 JS：RGB 任意通道 < 235 即视为有墨）
INK_THRESHOLD = 235
# 合并阈值（DETECT_SCALE 像素）
COL_GAP = 90
ROW_GAP = 180
# 行/列投影最小宽度阈值
MIN_BOX_W = 40
MIN_BAND_H = 40
GAP_TOLERANCE = 15  # 投影间断容忍行数


@dataclass
class Box:
    """DETECT_SCALE 像素坐标系下的矩形"""
    x: float
    y: float
    w: float
    h: float

    @property
    def cx(self) -> float:
        return self.x + self.w / 2

    @property
    def cy(self) -> float:
        return self.y + self.h / 2

    @property
    def x2(self) -> float:
        return self.x + self.w

    @property
    def y2(self) -> float:
        return self.y + self.h

    def to_dict(self) -> Dict[str, float]:
        return {"x": self.x, "y": self.y, "w": self.w, "h": self.h}


@dataclass
class Detection:
    """一次检测结果：raw 为未加 PAD 的核心框，padded 为加 PAD 后的外扩框"""
    raw: Box
    padded: Box


@dataclass
class Label:
    """单枚标签的源区域信息

    PyMuPDF / Pillow 坐标系：原点(0,0)在页面左上角，y 轴向下增长。

    居中策略：
        裁切框不是「检测框 + 外扩」，而是以标签自身的墨迹中心 (cx, cy) 为圆心、
        以统一标称尺寸 (crop_w × crop_h) 画一个矩形。这样所有标签：
        1) 尺寸完全相同（裁切框大小一致，缩放恒为 1.0）
        2) 内容完全居中（墨迹中心 == 裁切框中心 == 输出槽位中心）
    """
    src_idx: int   # 源 PDF 页索引
    cx: float      # 墨迹中心 x（源页面坐标系 pt）
    cy: float      # 墨迹中心 y（源页面坐标系 pt）
    core_w: float  # 该标签检测出的标称宽（pt，含 PAD_PX 留白）
    core_h: float  # 该标签检测出的标称高（pt，含 PAD_PX 留白）

    def crop_rect(self, crop_w: float, crop_h: float) -> fitz.Rect:
        """以墨迹中心为圆心生成统一尺寸的裁切框（fitz.Rect: x0,y0 左上 / x1,y1 右下）"""
        return fitz.Rect(
            self.cx - crop_w / 2,
            self.cy - crop_h / 2,
            self.cx + crop_w / 2,
            self.cy + crop_h / 2,
        )


# ================================================================
# 渲染与检测
# ================================================================

def render_page_pix(page: fitz.Page, scale: float = DETECT_SCALE) -> Image.Image:
    """渲染 PDF 页面为 PIL Image（用于检测）"""
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def _ink(img: Image.Image, x: int, y: int) -> bool:
    r, g, b = img.getpixel((x, y))
    return r < INK_THRESHOLD or g < INK_THRESHOLD or b < INK_THRESHOLD


def detect_stickers(page: fitz.Page) -> List[Detection]:
    """对单页 PDF 做投影法检测，返回候选标签列表

    每个 Detection 包含：
    - raw: 未加 PAD 的核心检测框（用于计算输出页面尺寸）
    - padded: 加 PAD 后的外扩框（用于后续裁切，防止切边）
    """
    img = render_page_pix(page)
    W, H = img.size

    # 行投影
    row_cnt = [0] * H
    for y in range(H):
        c = 0
        for x in range(0, W, 2):
            if _ink(img, x, y):
                c += 1
        row_cnt[y] = c
    row_on = lambda y: row_cnt[y] > 2

    # 抽取有墨的水平条带
    bands: List[Tuple[int, int]] = []
    s = -1
    gap = 0
    for y in range(H):
        if row_on(y):
            if s < 0:
                s = y
            gap = 0
        elif s >= 0:
            gap += 1
            if gap > GAP_TOLERANCE:
                if y - gap - s > MIN_BAND_H:
                    bands.append((s, y - gap))
                s = -1
    if s >= 0 and H - s > MIN_BAND_H:
        bands.append((s, H - 1))

    # 每个条带内做列投影
    boxes: List[Box] = []
    for y0, y1 in bands:
        col_cnt = [0] * W
        for x in range(W):
            for y in range(y0, y1, 2):
                if _ink(img, x, y):
                    col_cnt[x] += 1
        cs = -1
        for x in range(W + 1):
            on = x < W and col_cnt[x] > 1
            if on and cs < 0:
                cs = x
            if (not on or x == W) and cs >= 0:
                if x - cs > MIN_BOX_W:
                    boxes.append(Box(cs, y0, x - cs, y1 - y0))
                cs = -1

    merged = merge_boxes(boxes)

    # 生成 raw + padded 两种框
    result: List[Detection] = []
    for b in merged:
        nx = max(0, b.x - PAD_PX)
        ny = max(0, b.y - PAD_PX)
        nx2 = min(W, b.x2 + PAD_PX)
        ny2 = min(H, b.y2 + PAD_PX)
        raw = Box(b.x, b.y, b.w, b.h)
        padded = Box(nx, ny, nx2 - nx, ny2 - ny)
        result.append(Detection(raw=raw, padded=padded))
    return result


# ================================================================
# 框合并
# ================================================================

def _merge_group(group: List[Box]) -> Box:
    x = min(b.x for b in group)
    y = min(b.y for b in group)
    x2 = max(b.x2 for b in group)
    y2 = max(b.y2 for b in group)
    return Box(x, y, x2 - x, y2 - y)


def merge_boxes(boxes: List[Box], col_gap: float = COL_GAP, row_gap: float = ROW_GAP) -> List[Box]:
    """合并候选框：
    1) 按 x 中心聚列（容差 col_gap）
    2) 每列内按 y 聚行（容差 row_gap）
    3) 全局按 y 再合并一次，处理列聚类切开的边界
    """
    if not boxes:
        return []
    # 1. 按 x 中心聚列
    sorted_x = sorted(boxes, key=lambda b: b.cx)
    cols: List[Dict[str, Any]] = []
    for b in sorted_x:
        cx = b.cx
        target = None
        for c in cols:
            if abs(cx - c["sum"] / c["n"]) < col_gap:
                target = c
                break
        if target:
            target["boxes"].append(b)
            target["sum"] += cx
            target["n"] += 1
        else:
            cols.append({"boxes": [b], "sum": cx, "n": 1})

    # 2. 列内按 y 合并
    merged: List[Box] = []
    for col in cols:
        by_y = sorted(col["boxes"], key=lambda b: b.y)
        g = [by_y[0]]
        for i in range(1, len(by_y)):
            last = g[-1]
            dy = by_y[i].y - last.y2
            if dy < row_gap:
                g.append(by_y[i])
            else:
                merged.append(_merge_group(g))
                g = [by_y[i]]
        merged.append(_merge_group(g))

    # 3. 全局按 y 二次合并（处理同一标签被列聚类切开的边界）
    by_y = sorted(merged, key=lambda b: (b.y, b.x))
    if not by_y:
        return []
    final: List[Box] = []
    fg = [by_y[0]]
    for i in range(1, len(by_y)):
        last = fg[-1]
        dy = by_y[i].y - last.y2
        dx = abs(by_y[i].cx - last.cx)
        if dy < min(row_gap, 80) and dx < col_gap * 2:
            fg.append(by_y[i])
        else:
            final.append(_merge_group(fg))
            fg = [by_y[i]]
    final.append(_merge_group(fg))
    return final


def _canonical_size(values: List[float], tolerance: float = 1.15) -> float:
    """从一组检测尺寸中求「统一标称尺寸」

    思路：
    - 先取中位数，得到最可能的真实标签尺寸；
    - 再吸收与中位数接近的正常抖动（≤ tolerance 倍），取其中的最大值，
      保证没有任何一枚标签的内容被裁掉；
    - 明显偏大的框（如误把两枚标签合并成一枚）视为检测异常，不参与放大，
      否则会拖着所有标签一起变大。
    """
    if not values:
        return 0.0
    med = statistics.median(values)
    cands = [v for v in values if v <= med * tolerance]
    return max(cands) if cands else med


def _draw_label(
    out_page: "fitz.Page",
    src: "fitz.Document",
    lb: Label,
    crop: fitz.Rect,
    target: fitz.Rect,
) -> None:
    """把源页 crop 区域绘制到输出页 target 区域。

    裁切框可能越出源页边界（标签贴近页边时必然发生）。这里把裁切框与源页求交，
    只绘制落在页面内的部分，并按相同的位移量映射到目标区域的对应子矩形——
    越界部分保持白色，墨迹中心依然严格居中。
    """
    page_rect = src[lb.src_idx].rect
    clip = crop & page_rect
    if clip.width <= 0 or clip.height <= 0:
        return
    sub = fitz.Rect(
        target.x0 + (clip.x0 - crop.x0),
        target.y0 + (clip.y0 - crop.y0),
        target.x0 + (clip.x1 - crop.x0),
        target.y0 + (clip.y1 - crop.y0),
    )
    # pymupdf 1.24+ 改为静态方法：fitz.Page.show_pdf_page(page, rect, docsrc, pno, clip=...)
    fitz.Page.show_pdf_page(out_page, sub, src, lb.src_idx, clip=clip)


# ================================================================
# 主入口：PDF 重排
# ================================================================

def rearrange_pdf(
    input_path: str,
    output_path: str,
    per_page: int = 1,
    margin_mm: float = 0,
    padding_mm: float = 4,
) -> Dict[str, Any]:
    """主入口：检测 + 重排输出 PDF

    参数：
        input_path  : 源 PDF 路径
        output_path : 输出 PDF 路径
        per_page    : 每页输出几枚（默认 1）
        margin_mm   : 输出页面四周留白（毫米），仅影响页面大小
        padding_mm  : 检测框外扩余量（毫米），仅影响裁切范围，不影响页面大小

    返回：{'labels': 总标签数, 'pages': 输出页数, 'page_w_cm': 输出页宽cm, 'page_h_cm': 输出页高cm}
    """
    src = fitz.open(input_path)
    try:
        # 1. 检测所有页
        detections: List[List[Detection]] = []
        for p_idx in range(len(src)):
            detections.append(detect_stickers(src[p_idx]))

        # 2. 像素 → pt，构造 Label（只记录墨迹中心与标称尺寸）
        labels: List[Label] = []
        for p_idx, dets in enumerate(detections):
            for det in dets:
                r = det.raw
                labels.append(Label(
                    src_idx=p_idx,
                    cx=(r.x + r.w / 2) / DETECT_SCALE,
                    cy=(r.y + r.h / 2) / DETECT_SCALE,
                    core_w=det.padded.w / DETECT_SCALE,
                    core_h=det.padded.h / DETECT_SCALE,
                ))

        if not labels:
            raise ValueError("未检测到标签（请检查 PDF 内容或调整合并阈值）")

        # 3. 统一标称尺寸 → 统一裁切尺寸（所有标签完全一致）
        base_w = _canonical_size([l.core_w for l in labels])
        base_h = _canonical_size([l.core_h for l in labels])
        padding_pt = mm_to_pt(padding_mm)
        crop_w = base_w + padding_pt * 2
        crop_h = base_h + padding_pt * 2

        margin_pt = mm_to_pt(margin_mm)
        slot_gap = 6.0  # 多枚/页时槽位间固定间隔（pt）

        if per_page == 1:
            page_w = crop_w + margin_pt * 2
            page_h = crop_h + margin_pt * 2
        else:
            page_w = crop_w + margin_pt * 2
            page_h = per_page * crop_h + (per_page - 1) * slot_gap + margin_pt * 2

        n_pages = math.ceil(len(labels) / per_page)

        out = fitz.open()
        try:
            for _ in range(n_pages):
                out.new_page(width=page_w, height=page_h)
            # 注意：PyMuPDF 中缓存的 Page 引用会丢失 parent，必须用 out[i] 重新取

            for i, lb in enumerate(labels):
                # 裁切框：以该标签墨迹中心为圆心，尺寸恒为 crop_w × crop_h
                crop = lb.crop_rect(crop_w, crop_h)
                # 目标槽位：每枚标签尺寸一致，故缩放恒为 1.0，直接居中摆放
                if per_page == 1:
                    tx = margin_pt
                    ty = margin_pt
                else:
                    slot = i % per_page
                    tx = margin_pt
                    ty = margin_pt + slot * (crop_h + slot_gap)
                target = fitz.Rect(tx, ty, tx + crop_w, ty + crop_h)
                _draw_label(out[i // per_page], src, lb, crop, target)

            out.save(output_path)
        finally:
            out.close()
    finally:
        src.close()

    return {
        "labels": len(labels),
        "pages": n_pages,
        "page_w_cm": pt_to_cm(page_w),
        "page_h_cm": pt_to_cm(page_h),
        "label_w_cm": pt_to_cm(crop_w),
        "label_h_cm": pt_to_cm(crop_h),
    }