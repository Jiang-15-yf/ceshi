"""core.label_image - 生成带尺寸标注的标签参考图（嵌入 Excel）

原 HTML 用 Canvas 实现；本模块用 Pillow 等价实现，输出 PNG bytes。
"""
from __future__ import annotations

import io
from typing import Iterable, Optional

from PIL import Image, ImageDraw, ImageFont

# 标注色与原始 JS 版保持一致
ANNOTATION_COLOR = "#dc2626"


def _load_font(size: int) -> ImageFont.ImageFont:
    """尝试加载常见中文字体，失败时回退默认"""
    candidates = [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simsun.ttc",
        "/System/Library/Fonts/PingFang.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _draw_size_annotation(d: ImageDraw.ImageDraw, W: int, H: int, w_cm: float, h_cm: float) -> None:
    """在画布 d 上叠加尺寸标注（边框 + 底部宽文字 + 中心大文字）"""
    pad = 12
    # 外框
    d.rectangle([pad, pad, W - pad, H - pad], outline=ANNOTATION_COLOR, width=2)
    # 底部宽标注（箭头线 + 文字）
    arrow_y = H - pad - 6
    d.line([(pad + 10, arrow_y), (W - pad - 10, arrow_y)], fill=ANNOTATION_COLOR, width=2)
    d.line([(pad + 10, H - pad - 10), (pad + 10, H - pad - 2)], fill=ANNOTATION_COLOR, width=2)
    d.line([(W - pad - 10, H - pad - 10), (W - pad - 10, H - pad - 2)], fill=ANNOTATION_COLOR, width=2)
    font_big = _load_font(18)
    txt = f"{w_cm:.1f} cm"
    bbox = d.textbbox((0, 0), txt, font=font_big)
    d.text(((W - bbox[2]) / 2, H - pad - bbox[3] - 14), txt, fill=ANNOTATION_COLOR, font=font_big)
    # 中心大文字
    font_center = _load_font(22)
    txt2 = f"{w_cm:.1f} × {h_cm:.1f} cm"
    bbox = d.textbbox((0, 0), txt2, font=font_center)
    d.text(((W - bbox[2]) / 2, (H - bbox[3]) / 2), txt2, fill=ANNOTATION_COLOR, font=font_center)


def make_label_diagram(w_cm: float, h_cm: float) -> bytes:
    """生成空白示意图（无参考图时使用），返回 PNG bytes"""
    W = max(200, round(w_cm * 90))
    H = max(200, round(h_cm * 90))
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    _draw_size_annotation(d, W, H, w_cm, h_cm)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def make_annotated_from_bytes(image_bytes: bytes, w_cm: float, h_cm: float) -> bytes:
    """将用户上传的标签图按目标宽高比裁剪填充并返回 PNG bytes

    注意：用户上传的参考图通常已自带尺寸标注，因此这里不再叠加额外标注，
    仅做等比缩放/填充，保证嵌入 Excel 后不变形、不裁切主要内容。
    """
    src = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    target_ratio = w_cm / h_cm
    # 计算目标画布尺寸（保持原图比例与目标比例对齐）
    if src.width / src.height > target_ratio:
        H = round(src.width / target_ratio)
        W = src.width
    else:
        W = round(src.height * target_ratio)
        H = src.height
    # 等比缩放并居中，确保原图内容完整不被裁切
    # 背景白色，避免 target_ratio 与原图比例不一致时露黑边
    draw_ratio = min(W / src.width, H / src.height)
    bg = Image.new("RGB", (W, H), "white")
    dw, dh = src.width * draw_ratio, src.height * draw_ratio
    bg.paste(src, (round((W - dw) / 2), round((H - dh) / 2)))
    buf = io.BytesIO()
    bg.save(buf, format="PNG")
    return buf.getvalue()