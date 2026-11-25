# -*- coding: utf-8 -*-
"""生成 PPT 工具箱 Logo：露着半截 PPT 的工具箱 + 清风字样。"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont

S = 4                      # 超采样倍数
W = 1024                   # 设计画布尺寸
BG = (18, 185, 129)        # 品牌绿
MINT = (125, 235, 192)     # 薄荷
DARK = (10, 49, 40)        # 箱体深墨绿
WHITE = (255, 255, 255)

FONT_DIR = r"C:\Windows\Fonts"


def font(name, size):
    for cand in (name, "msyhbd.ttc", "msyh.ttc", "simhei.ttf"):
        p = os.path.join(FONT_DIR, cand)
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


def rr(d, box, r, fill=None, outline=None, width=1):
    d.rounded_rectangle([box[0] * S, box[1] * S, box[2] * S, box[3] * S],
                        radius=r * S, fill=fill, outline=outline,
                        width=width * S)


def sheet(size, radius, with_bars=True):
    """一张“PPT 页”：白底 + 标题条 + 柱状图，返回未旋转的透明图。"""
    w, h = size
    img = Image.new("RGBA", (w * S, h * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, w * S - 1, h * S - 1], radius=radius * S,
                        fill=WHITE, outline=(226, 240, 233), width=2 * S)
    # 标题条
    d.rounded_rectangle([w * 0.10 * S, h * 0.13 * S,
                         w * 0.56 * S, h * 0.19 * S],
                        radius=6 * S, fill=(206, 232, 220))
    if with_bars:
        base = h * 0.60
        bw, gap = w * 0.135, w * 0.085
        x = w * 0.13
        for i, c in enumerate((BG, (52, 205, 150), MINT)):
            bh = h * (0.16 + 0.11 * i)
            d.rounded_rectangle([x * S, (base - bh) * S,
                                 (x + bw) * S, base * S],
                                radius=6 * S, fill=c)
            x += bw + gap
        # 基线
        d.rounded_rectangle([w * 0.10 * S, base * S,
                             w * 0.93 * S, (base + 5) * S],
                            radius=3 * S, fill=(226, 240, 233))
    return img


def build(with_text=True):
    canvas = Image.new("RGBA", (W * S, W * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(canvas)

    # 底板
    rr(d, (20, 20, W - 20, W - 20), 228, fill=BG)

    # 后面一页（只露个边，不做细节）
    back = sheet((280, 236), 24, with_bars=False)
    back = back.rotate(7, expand=True, resample=Image.BICUBIC)
    canvas.alpha_composite(back, (int(430 * S), int(286 * S)))

    # 前面一页（半截露在箱口外）
    front = sheet((420, 268), 26)
    front = front.rotate(-7, expand=True, resample=Image.BICUBIC)
    canvas.alpha_composite(front, (int(238 * S), int(268 * S)))

    # 箱体
    rr(d, (146, 552, 878, 850), 42, fill=DARK)
    # 箱盖
    rr(d, (122, 480, 902, 582), 46, fill=MINT)
    # 提手（箱盖正面的拉手：横把 + 一条透缝）
    rr(d, (398, 542, 626, 606), 32, fill=DARK)
    rr(d, (450, 564, 574, 586), 11, fill=MINT)

    if with_text:
        f = font("msyhbd.ttc", 180 * S)
        txt = "清风"
        bb = d.textbbox((0, 0), txt, font=f)
        tw, th = bb[2] - bb[0], bb[3] - bb[1]
        d.text((W * S / 2 - tw / 2 - bb[0], 716 * S - th / 2 - bb[1]),
               txt, font=f, fill=WHITE)

    return canvas.resize((2048, 2048), Image.LANCZOS)


def _png_bytes(img, size):
    import io
    buf = io.BytesIO()
    img.resize((size, size), Image.LANCZOS).save(buf, format="PNG")
    return buf.getvalue()


def write_ico(path, big_img, small_img):
    """自己拼 ICO：≥48px 用带字版，≤32px 用无字简化版，避免小图标糊成一团。"""
    sizes = [(256, 256, big_img), (128, 128, big_img), (96, 96, big_img),
             (64, 64, big_img),
             (48, 48, small_img), (32, 32, small_img),
             (24, 24, small_img), (16, 16, small_img)]
    blobs = [_png_bytes(img, s[0]) for s in sizes for img in [s[2]]]
    head = bytearray(b"\x00\x00\x01\x00" + len(sizes).to_bytes(2, "little"))
    offset = 6 + 16 * len(sizes)
    entries = bytearray()
    for (w, h, _), blob in zip(sizes, blobs):
        entries += bytes([0 if w == 256 else w, 0 if h == 256 else h, 0, 0])
        entries += (1).to_bytes(2, "little") + (32).to_bytes(2, "little")
        entries += len(blob).to_bytes(4, "little")
        entries += offset.to_bytes(4, "little")
        offset += len(blob)
    with open(path, "wb") as f:
        f.write(bytes(head) + bytes(entries) + b"".join(blobs))
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "logo_preview.png"
    big = build(True)
    small = build(False)
    big.save(out)
    big.resize((512, 512), Image.LANCZOS).save("logo_512.png")
    big.resize((160, 160), Image.LANCZOS).save("logo_160.png")
    small.resize((512, 512), Image.LANCZOS).save("assets/logo_mark.png")
    write_ico("app.ico", big, small)
    print("saved", out, "app.ico", "logo_512.png", "logo_160.png")
