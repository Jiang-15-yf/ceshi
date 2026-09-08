# -*- coding: utf-8 -*-
"""main - Tkinter 桌面 GUI 入口（极简风格 v1.3）

布局：顶栏 + 左侧导航 + 内容区 + 底栏。
- Excel 标签：选择表 → 参数 → 预览（电子表格画布，与导出完全一致）→ 导出
- PDF 重排：选择 PDF → 解析 → 参数 → 生成

业务逻辑全部在 core/* 中，本文件只做 UI 编排。
"""
from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont
from tkinter import Tk, StringVar, DoubleVar, IntVar, END, DISABLED, NORMAL, Button, Frame, Label
from tkinter import Entry as TkEntry
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk

from core.pdf_rearrange import rearrange_pdf, detect_stickers, DETECT_SCALE
from core.excel_pricecard import parse_table, build_pricecard, export_xlsx, build_sheet_model
from core.sheet_view import SheetView


APP_TITLE = "价格牌处理工具"
APP_VERSION = "1.3"

# ================================================================
# 资源路径（兼容开发目录与 PyInstaller 单文件打包）
# ================================================================
def _resource_path(rel: str) -> str:
    """返回资源文件绝对路径。PyInstaller 单文件运行时资源在 _MEIPASS。"""
    if hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent.resolve()
    return str(base / rel)


# ================================================================
# 配色（极简 · 近黑 + 发丝级灰线 + 大量留白）
# ================================================================
INK = "#17181C"          # 主文字 / 主按钮
INK_2 = "#5C616B"        # 次要文字
INK_3 = "#9AA0A8"        # 弱文字
LINE = "#ECECF0"         # 发丝边框
LINE_2 = "#F5F5F7"       # 浅填充
BG = "#FFFFFF"           # 内容背景
BG_APP = "#F6F6F8"       # 应用背景
SIDEBAR = "#FAFAFB"      # 侧栏背景
ACCENT = "#2F6BFF"       # 仅用于激活态指示的克制强调
SELECT = "#EFF1F4"       # 选中/悬停浅底

FONT_TITLE = ("Microsoft YaHei", 14, "bold")
FONT_SUB = ("Microsoft YaHei", 10)
FONT_NAV = ("Microsoft YaHei", 11)
FONT_LABEL = ("Microsoft YaHei", 10)
FONT_BOLD = ("Microsoft YaHei", 10, "bold")
FONT_MONO = ("Consolas", 10)
FONT_HINT = ("Microsoft YaHei", 9)
FONT_BUTTON = ("Microsoft YaHei", 10)

CORNER = 8  # 圆角半径（px）


# ================================================================
# 圆角控件
# ================================================================

class RoundedButton(tk.Canvas):
    """圆角按钮：primary=近黑填充；secondary=白底发丝边框。

    用 Canvas 绘制圆角矩形 + 居中文字，规避 tk.Button 无法圆角的问题。
    """

    def __init__(self, parent, text="", kind="primary", command=None, state="normal",
                 padx=18, pady=8, font=FONT_BUTTON, bg=BG, **kw):
        super().__init__(parent, bd=0, highlightthickness=0, bg=bg, **kw)
        self._text = text
        self._kind = kind
        self._command = command
        self._state = state
        self._padx = padx
        self._pady = pady
        self._font = font
        self._hover = False
        self._items = []
        f = tkfont.Font(family=font[0], size=font[1],
                        weight=font[2] if len(font) > 2 else "normal")
        w = f.measure(text) + padx * 2
        h = f.metrics("linespace") + pady * 2
        self._bw, self._bh = max(int(w), 44), int(h)
        self.config(width=self._bw, height=self._bh)
        self._draw()
        self.bind("<Enter>", lambda e: self._enter())
        self.bind("<Leave>", lambda e: self._leave())
        self.bind("<Button-1>", lambda e: self._press())
        self.bind("<ButtonRelease-1>", lambda e: self._release())
        self.config(cursor="hand2" if state == "normal" else "arrow")

    def _colors(self):
        if self._state == DISABLED:
            return {"fill": "#EDEEF0", "outline": "#EDEEF0", "fg": "#C7C9CE"}
        if self._kind == "primary":
            fill = "#000000" if self._hover else INK
            return {"fill": fill, "outline": fill, "fg": "#FFFFFF"}
        fill = SELECT if self._hover else BG
        return {"fill": fill, "outline": LINE, "fg": INK}

    def _draw(self):
        for it in self._items:
            self.delete(it)
        self._items = []
        c = self._colors()
        self._items.extend(self._round_rect(1, 1, self._bw - 2, self._bh - 2, CORNER,
                                            fill=c["fill"], outline=c["outline"], width=1))
        self._items.append(self.create_text(self._bw / 2, self._bh / 2, text=self._text,
                                            fill=c["fg"], font=self._font, anchor="center"))

    def _round_rect(self, x1, y1, x2, y2, r, fill=BG, outline=LINE, width=1):
        """绘制圆角矩形：先填充（外描边与填充同色，避免内部出现横线），再描边。"""
        its = []
        kw_fill = dict(fill=fill, outline=fill, width=1)
        its.append(self.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="pieslice", **kw_fill))
        its.append(self.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="pieslice", **kw_fill))
        its.append(self.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="pieslice", **kw_fill))
        its.append(self.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="pieslice", **kw_fill))
        its.append(self.create_rectangle(x1 + r, y1, x2 - r, y2, **kw_fill))
        its.append(self.create_rectangle(x1, y1 + r, x2, y2 - r, **kw_fill))
        if outline and width > 0:
            kw_arc = dict(fill="", outline=outline, width=width)
            kw_line = dict(fill=outline, width=width)
            its.append(self.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="arc", **kw_arc))
            its.append(self.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="arc", **kw_arc))
            its.append(self.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="arc", **kw_arc))
            its.append(self.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="arc", **kw_arc))
            its.append(self.create_line(x1 + r, y1, x2 - r, y1, **kw_line))
            its.append(self.create_line(x1 + r, y2, x2 - r, y2, **kw_line))
            its.append(self.create_line(x1, y1 + r, x1, y2 - r, **kw_line))
            its.append(self.create_line(x2, y1 + r, x2, y2 - r, **kw_line))
        return its

    def _enter(self):
        if self._state == DISABLED:
            return
        self._hover = True
        self._draw()

    def _leave(self):
        self._hover = False
        self._draw()

    def _press(self):
        if self._state == DISABLED or not self._command:
            return

    def _release(self):
        if self._state == DISABLED or not self._command:
            return
        self._command()

    def set_state(self, state):
        self._state = state
        self.config(cursor="hand2" if state == "normal" else "arrow")
        self._draw()


class RoundedField(tk.Frame):
    """圆角输入框：Canvas 绘制圆角边框，内部 Entry 无边框嵌入。"""

    def __init__(self, parent, var, width=9, font=FONT_LABEL, bg=BG, **kw):
        super().__init__(parent, bg=bg, **kw)
        self._bg = bg
        self._canvas = tk.Canvas(self, bg=bg, bd=0, highlightthickness=0)
        self._canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._entry = TkEntry(self, textvariable=var, width=width, font=font,
                              bg=bg, fg=INK, relief="flat", bd=0,
                              insertbackground=INK, highlightthickness=0)
        self._entry.pack(padx=10, pady=5)
        self._focus = False
        self._entry.bind("<FocusIn>", lambda e: self._redraw(True))
        self._entry.bind("<FocusOut>", lambda e: self._redraw(False))
        self._canvas.bind("<Configure>", lambda e: self._redraw(self._focus))

    def _redraw(self, focus):
        self._focus = focus
        self._canvas.delete("all")
        w = self._canvas.winfo_width()
        h = self._canvas.winfo_height()
        if w < 4 or h < 4:
            return
        color = ACCENT if focus else LINE
        r = CORNER
        self._round_rect(1, 1, w - 2, h - 2, r, fill=self._bg,
                         outline=color, width=1 if not focus else 1.5)

    def _round_rect(self, x1, y1, x2, y2, r, fill=BG, outline=LINE, width=1):
        c = self._canvas
        kw_fill = dict(fill=fill, outline=fill, width=1)
        c.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="pieslice", **kw_fill)
        c.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="pieslice", **kw_fill)
        c.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="pieslice", **kw_fill)
        c.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="pieslice", **kw_fill)
        c.create_rectangle(x1 + r, y1, x2 - r, y2, **kw_fill)
        c.create_rectangle(x1, y1 + r, x2, y2 - r, **kw_fill)
        if outline and width > 0:
            kw_arc = dict(fill="", outline=outline, width=width)
            kw_line = dict(fill=outline, width=width)
            c.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="arc", **kw_arc)
            c.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="arc", **kw_arc)
            c.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="arc", **kw_arc)
            c.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="arc", **kw_arc)
            c.create_line(x1 + r, y1, x2 - r, y1, **kw_line)
            c.create_line(x1 + r, y2, x2 - r, y2, **kw_line)
            c.create_line(x1, y1 + r, x1, y2 - r, **kw_line)
            c.create_line(x2, y1 + r, x2, y2 - r, **kw_line)


# ================================================================
# 通用工具
# ================================================================

def _set_status(text_widget, msg: str) -> None:
    text_widget.config(state=NORMAL)
    text_widget.delete("1.0", END)
    if msg:
        text_widget.insert("1.0", msg)
    text_widget.config(state=DISABLED)


def _btn(parent, text: str, kind: str = "primary", command=None, state: str = "normal", width=None) -> RoundedButton:
    """圆角按钮：primary=近黑填充；secondary=白底发丝边框。"""
    return RoundedButton(parent, text=text, kind=kind, command=command, state=state)


def set_btn_state(btn, state: str) -> None:
    if hasattr(btn, "set_state"):
        btn.set_state(state)
    else:
        btn.config(state=state)


def _field(parent, label: str, var, hint: str = "", width: int = 9) -> Frame:
    """一行参数：标签 + 圆角输入框 + 提示"""
    row = Frame(parent, bg=BG)
    Label(row, text=label, font=FONT_LABEL, bg=BG, fg=INK_2).pack(side="left")
    field = RoundedField(row, var, width=width)
    field.pack(side="left", padx=8)
    if hint:
        Label(row, text=hint, font=FONT_HINT, bg=BG, fg=INK_3).pack(side="left")
    return row


# ================================================================
# 主窗口
# ================================================================

class App:
    def __init__(self, root: Tk) -> None:
        self.root = root
        root.title(f"{APP_TITLE}  v{APP_VERSION}")
        root.geometry("1120x860")
        root.minsize(940, 720)
        root.configure(bg=BG_APP)
        try:
            ttk.Style().theme_use("clam")
        except Exception:
            pass

        self._build_topbar()
        self._build_sidebar()
        self._build_content_host()
        self._build_statusbar()

        self.current = None
        self._nav_to("excel")

    # ----------------------------------------------------------
    # 框架
    # ----------------------------------------------------------
    def _build_topbar(self) -> None:
        bar = Frame(self.root, bg=BG, height=60)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)
        inner = Frame(bar, bg=BG)
        inner.pack(fill="x", expand=True, padx=24)
        # 标志：Logo（价格签形状）
        try:
            self.logo_img = ImageTk.PhotoImage(
                Image.open(_resource_path("assets/logo_22.png")).convert("RGBA")
            )
            logo_lbl = Label(inner, image=self.logo_img, bg=BG)
            logo_lbl.pack(side="left", padx=(0, 12))
        except Exception as e:
            # 兜底：如果 Logo 资源缺失，仍用近黑方块保证界面不崩
            mark = Frame(inner, bg=INK, width=22, height=22)
            mark.pack(side="left", padx=(0, 12))
            print(f"[warn] Logo 加载失败: {e}")
        # 标题（pack 自动计算宽度，避免 place 导致的截断）
        Label(inner, text=APP_TITLE, font=FONT_TITLE, bg=BG, fg=INK).pack(side="left")
        # 版本靠右
        Label(inner, text=f"v{APP_VERSION}", font=FONT_HINT, bg=BG, fg=INK_3).pack(side="right")
        # 底线
        sep = Frame(self.root, bg=LINE, height=1)
        sep.pack(fill="x", side="top")

    def _build_sidebar(self) -> None:
        side = Frame(self.root, bg=SIDEBAR, width=245)
        side.pack(fill="y", side="left")
        side.pack_propagate(False)
        Label(side, text="功能", font=FONT_HINT, bg=SIDEBAR, fg=INK_3).pack(anchor="w", padx=20, pady=(22, 10))

        self.nav_items = {}
        self._nav_row(side, "excel", "①  Excel 生产数量统计转化")
        self._nav_row(side, "pdf", "②  PDF 价格签重排")

    def _nav_row(self, side, key, text) -> None:
        item = Frame(side, bg=SIDEBAR)
        item.pack(fill="x", pady=2)
        bar = Frame(item, bg=SIDEBAR, width=3)
        bar.pack(side="left", fill="y", padx=(16, 0))
        btn = Button(item, text=text, font=FONT_NAV, bg=SIDEBAR, fg=INK_2,
                     bd=0, relief="flat", anchor="w", padx=10, pady=8,
                     activebackground=SIDEBAR, activeforeground=INK,
                     command=lambda k=key: self._nav_to(k))
        btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.nav_items[key] = (item, bar, btn)

    def _build_content_host(self) -> None:
        self.host = Frame(self.root, bg=BG)
        self.host.pack(fill="both", expand=True, side="left")

        self.excel_tab = Frame(self.host, bg=BG)
        self.pdf_tab = Frame(self.host, bg=BG)
        self._build_excel_tab(self.excel_tab)
        self._build_pdf_tab(self.pdf_tab)

    def _build_statusbar(self) -> None:
        bar = Frame(self.root, bg=BG, height=30)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        sep = Frame(self.root, bg=LINE, height=1)
        sep.pack(fill="x", side="bottom")
        self.status_var = StringVar(value="就绪")
        Label(bar, textvariable=self.status_var, font=FONT_HINT, bg=BG, fg=INK_2).pack(side="left", padx=24)

    def _nav_to(self, key: str) -> None:
        for k, (item, bar, btn) in self.nav_items.items():
            active = (k == key)
            btn.config(bg=BG if active else SIDEBAR, fg=INK if active else INK_2)
            bar.config(bg=INK if active else SIDEBAR)
        if self.current and self.current != key:
            getattr(self, f"{self.current}_tab").pack_forget()
        self.current = key
        tab = getattr(self, f"{key}_tab")
        tab.pack(fill="both", expand=True, padx=0, pady=0)
        self.status_var.set("就绪")

    # ----------------------------------------------------------
    # 区块工具
    # ----------------------------------------------------------
    def _section(self, parent, title: str, subtitle: str = "") -> Frame:
        """内容区的一个区块：标题 + 分隔线 + 内容 Frame"""
        wrap = Frame(parent, bg=BG)
        wrap.pack(fill="x", padx=28, pady=(8, 4))
        head = Frame(wrap, bg=BG)
        head.pack(fill="x")
        Label(head, text=title, font=("Microsoft YaHei", 12, "bold"), bg=BG, fg=INK).pack(side="left")
        if subtitle:
            Label(head, text=subtitle, font=FONT_HINT, bg=BG, fg=INK_3).pack(side="left", padx=10)
        Line = Frame(wrap, bg=LINE, height=1)
        Line.pack(fill="x", pady=(6, 10))
        return wrap

    # ----------------------------------------------------------
    # Excel 标签 Tab
    # ----------------------------------------------------------
    def _build_excel_tab(self, parent: Frame) -> None:
        # 输入：文件 + 参考图
        sec = self._section(parent, "输入", "Excel / CSV / XLS · 可选标签参考图")
        frow = Frame(sec, bg=BG)
        frow.pack(fill="x", pady=(0, 8))
        self.ex_file_path = StringVar()
        _btn(frow, "选择文件", "secondary", self._ex_pick_file).pack(side="left")
        self.ex_path_lbl = Label(frow, text="未选择文件", font=FONT_LABEL, bg=BG, fg=INK_3, anchor="w")
        self.ex_path_lbl.pack(side="left", padx=14, fill="x", expand=True)

        irow = Frame(sec, bg=BG)
        irow.pack(fill="x")
        self.ex_img_path = StringVar()
        _btn(irow, "选择参考图", "secondary", self._ex_pick_image).pack(side="left")
        self.ex_img_lbl = Label(irow, text="未选择（使用默认标签图）", font=FONT_LABEL, bg=BG, fg=INK_3, anchor="w")
        self.ex_img_lbl.pack(side="left", padx=14, fill="x", expand=True)

        # 参数
        sec2 = self._section(parent, "参数")
        self.ex_rate = DoubleVar(value=4.0)
        self.tag_w = DoubleVar(value=4.5)
        self.tag_h = DoubleVar(value=9.0)
        prow1 = Frame(sec2, bg=BG)
        prow1.pack(fill="x", pady=(0, 10))
        _field(prow1, "生产数上浮 %", self.ex_rate, "(0~20)").pack(side="left")
        prow2 = Frame(sec2, bg=BG)
        prow2.pack(fill="x")
        _field(prow2, "标签宽 cm", self.tag_w, "(推荐 4.5)").pack(side="left", padx=(0, 28))
        _field(prow2, "标签高 cm", self.tag_h, "(推荐 9.0)").pack(side="left")

        # 操作（右对齐：预览 / 导出）
        arow = Frame(parent, bg=BG)
        arow.pack(fill="x", padx=28, pady=12)
        self.ex_go_btn = _btn(arow, "导出 .xlsx", "primary", self._ex_generate, state=DISABLED)
        self.ex_go_btn.pack(side="right")
        _btn(arow, "预览", "secondary", self._ex_preview).pack(side="right", padx=(0, 10))

        # 预览区
        sec4 = self._section(parent, "工作表预览", "与导出的 .xlsx 完全一致")
        prev_wrap = Frame(parent, bg=BG)
        prev_wrap.pack(fill="both", expand=True, padx=28, pady=(0, 18))
        border = Frame(prev_wrap, bg=LINE, bd=1, relief="solid")
        border.pack(fill="both", expand=True)
        self.ex_sheet = SheetView(border)
        self.ex_sheet.pack(fill="both", expand=True, padx=1, pady=1)

        self.ex_status = StringVar(value="")
        self.ex_parsed = None
        self.ex_ref_image_bytes = None

    def _ex_pick_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel/CSV 文件",
            filetypes=[("Excel/CSV", "*.xls *.xlsx *.csv"), ("All", "*.*")],
        )
        if not path:
            return
        self.ex_file_path.set(path)
        self.ex_path_lbl.config(text=os.path.basename(path), fg=INK)
        try:
            self.ex_parsed = parse_table(path)
            set_btn_state(self.ex_go_btn, NORMAL)
            n = len(self.ex_parsed["rows"])
            sizes = "、".join(self.ex_parsed["size_headers"])
            self.status_var.set(f"已读取：{n} 行 · 尺码 [{sizes}]")
        except Exception as e:
            self.ex_parsed = None
            set_btn_state(self.ex_go_btn, DISABLED)
            self.status_var.set("读取失败")
            messagebox.showerror("读取失败", str(e))

    def _ex_pick_image(self) -> None:
        path = filedialog.askopenfilename(
            title="选择标签参考图",
            filetypes=[("Image", "*.png *.jpg *.jpeg *.bmp"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "rb") as f:
                self.ex_ref_image_bytes = f.read()
            self.ex_img_path.set(path)
            self.ex_img_lbl.config(text=os.path.basename(path), fg=INK)
            self.status_var.set(f"已载入参考图：{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("载入失败", str(e))

    def _ex_preview(self) -> None:
        if not self.ex_parsed:
            messagebox.showwarning("提示", "请先选择文件")
            return
        try:
            built = build_pricecard(self.ex_parsed, self.ex_rate.get())
            model = build_sheet_model(
                self.ex_parsed, built,
                tag_w_cm=self.tag_w.get(), tag_h_cm=self.tag_h.get(),
                ref_image_bytes=self.ex_ref_image_bytes,
            )
            self.ex_sheet.set_model(model)
            body_n = len(built["body"]); blank_n = len(built["blank_out"])
            self.status_var.set(f"预览：主体 {body_n} 行 + 空白 {blank_n} 行 + 合计 1 行")
            self.ex_model = model
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("预览失败", str(e))

    def _ex_generate(self) -> None:
        if not self.ex_parsed:
            return
        style0 = self.ex_parsed["rows"][0]["style"] if self.ex_parsed["rows"] else "特殊价格牌"
        default_name = f"{style0 or '特殊价格牌'}-特殊价格牌.xlsx"
        out_path = filedialog.asksaveasfilename(
            title="保存为",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=default_name,
        )
        if not out_path:
            return
        try:
            built = build_pricecard(self.ex_parsed, self.ex_rate.get())
            model = build_sheet_model(
                self.ex_parsed, built,
                tag_w_cm=self.tag_w.get(), tag_h_cm=self.tag_h.get(),
                ref_image_bytes=self.ex_ref_image_bytes,
            )
            export_xlsx(self.ex_parsed, built, out_path,
                        tag_w_cm=self.tag_w.get(), tag_h_cm=self.tag_h.get(),
                        ref_image_bytes=self.ex_ref_image_bytes)
            self.ex_sheet.set_model(model)
            self.status_var.set(f"已导出：{out_path}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("生成失败", str(e))

    # ----------------------------------------------------------
    # PDF 标签 Tab
    # ----------------------------------------------------------
    def _build_pdf_tab(self, parent: Frame) -> None:
        sec = self._section(parent, "输入", "每页含多个价格签的 PDF")
        frow = Frame(sec, bg=BG)
        frow.pack(fill="x")
        self.pd_file_path = StringVar()
        _btn(frow, "选择 PDF", "secondary", self._pd_pick_file).pack(side="left")
        self.pd_path_lbl = Label(frow, text="未选择文件", font=FONT_LABEL, bg=BG, fg=INK_3, anchor="w")
        self.pd_path_lbl.pack(side="left", padx=14, fill="x", expand=True)
        self.pd_analyze_btn = _btn(frow, "解析并检测", "secondary", self._pd_analyze, state=DISABLED)
        self.pd_analyze_btn.pack(side="left", padx=(10, 0))

        sec2 = self._section(parent, "检测结果", "")
        self.pd_info = Label(sec2, text="请选择 PDF 并点击「解析并检测」", font=FONT_LABEL, bg=BG, fg=INK_2, anchor="w", justify="left")
        self.pd_info.pack(fill="x")

        sec3 = self._section(parent, "输出参数")
        prow = Frame(sec3, bg=BG)
        prow.pack(fill="x")
        self.pd_per_page = IntVar(value=1)
        self.pd_margin = DoubleVar(value=0)
        self.pd_padding = DoubleVar(value=0)
        _field(prow, "每页标签数", self.pd_per_page).pack(side="left", padx=(0, 28))
        _field(prow, "页边距 mm", self.pd_margin).pack(side="left", padx=(0, 28))
        _field(prow, "防切边 mm", self.pd_padding).pack(side="left")

        arow = Frame(parent, bg=BG)
        arow.pack(fill="x", padx=28, pady=12)
        _btn(arow, "生成重排 PDF", "primary", self._pd_generate).pack(side="right")

        sec4 = self._section(parent, "输出信息", "")
        self.pd_status = ScrolledText(sec4, height=5, state=DISABLED, font=FONT_MONO,
                                     bd=0, bg=LINE_2, fg=INK, relief="flat", padx=10, pady=8)
        self.pd_status.pack(fill="x")

        self.pd_doc = None
        self.pd_detections = []

    def _pd_pick_file(self) -> None:
        path = filedialog.askopenfilename(title="选择 PDF 文件", filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self.pd_file_path.set(path)
        self.pd_path_lbl.config(text=os.path.basename(path), fg=INK)
        set_btn_state(self.pd_analyze_btn, NORMAL)
        _set_status(self.pd_status, "已选择，请点击「解析并检测」")
        self.status_var.set(f"已选择：{os.path.basename(path)}")

    def _pd_analyze(self) -> None:
        path = self.pd_file_path.get()
        if not path:
            return
        try:
            import fitz
            doc = fitz.open(path)
            self.pd_doc = doc
            self.pd_detections = []
            total = 0
            lines = []
            for i in range(len(doc)):
                boxes = detect_stickers(doc[i])
                self.pd_detections.append(boxes)
                total += len(boxes)
                lines.append(f"第 {i+1} 页：{len(boxes)} 枚")
            sizes = set()
            for boxes in self.pd_detections:
                for det in boxes:
                    w = det.padded.w / DETECT_SCALE
                    h = det.padded.h / DETECT_SCALE
                    sizes.add((round(w / 72 * 2.54, 2), round(h / 72 * 2.54, 2)))
            sizes_str = ", ".join([f"{w}×{h}cm" for w, h in sorted(sizes)])
            self.pd_info.config(
                text=f"共 {len(doc)} 页 · 检测到 {total} 枚标签\n检测尺寸：{sizes_str}\n" + "\n".join(lines),
                fg=INK)
            _set_status(self.pd_status, "检测完成，可以生成。每枚标签将严格居中、尺寸完全一致。")
            self.status_var.set(f"检测完成：{total} 枚标签")
        except Exception as e:
            _set_status(self.pd_status, f"解析失败：{e}")
            self.status_var.set("解析失败")

    def _pd_generate(self) -> None:
        path = self.pd_file_path.get()
        if not path:
            messagebox.showwarning("提示", "请先选择 PDF 文件")
            return
        out_path = filedialog.asksaveasfilename(
            title="保存为",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=Path(path).stem + "_价格签重排.pdf",
        )
        if not out_path:
            return
        try:
            result = rearrange_pdf(
                input_path=path,
                output_path=out_path,
                per_page=self.pd_per_page.get(),
                margin_mm=self.pd_margin.get(),
                padding_mm=self.pd_padding.get(),
            )
            msg = (
                f"已生成：{os.path.basename(out_path)}\n"
                f"标签数 {result['labels']} · 输出 {result['pages']} 页\n"
                f"单枚标签尺寸：{result['label_w_cm']:.2f} × {result['label_h_cm']:.2f} cm\n"
                f"输出页面尺寸：{result['page_w_cm']:.2f} × {result['page_h_cm']:.2f} cm\n"
                f"所有标签严格居中、尺寸完全一致。"
            )
            _set_status(self.pd_status, msg)
            self.status_var.set(f"已生成：{os.path.basename(out_path)}")
        except Exception as e:
            _set_status(self.pd_status, f"生成失败：{e}")
            self.status_var.set("生成失败")


# ================================================================
# 入口
# ================================================================

def main() -> None:
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
