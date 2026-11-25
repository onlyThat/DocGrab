"""生成 docgrab.ico：一个"文档 + 提取箭头"的简洁图标。"""
from __future__ import annotations

import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), "docgrab.ico")


def draw(size: int) -> Image.Image:
    s = size
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = int(s * 0.10)
    # 蓝色圆角背景
    d.rounded_rectangle([pad, pad, s - pad, s - pad], radius=int(s * 0.18),
                        fill=(58, 123, 213, 255))
    # 白色文档
    docx0 = int(s * 0.26)
    docy0 = int(s * 0.16)
    docx1 = int(s * 0.62)
    docy1 = int(s * 0.70)
    d.rounded_rectangle([docx0, docy0, docx1, docy1], radius=int(s * 0.03),
                        fill=(255, 255, 255, 255))
    # 文档上的横线
    line_col = (58, 123, 213, 255)
    lx0 = docx0 + int(s * 0.05)
    lx1 = docx1 - int(s * 0.05)
    for i in range(4):
        y = docy0 + int(s * 0.10) + i * int(s * 0.09)
        d.line([lx0, y, lx1, y], fill=line_col, width=max(1, int(s * 0.018)))
    # 绿色提取箭头
    ax = int(s * 0.78)
    ay0 = int(s * 0.30)
    ay1 = int(s * 0.80)
    d.line([ax, ay0, ax, ay1], fill=(46, 170, 90, 255), width=max(2, int(s * 0.05)))
    arrow = [(ax - int(s * 0.09), ay1 - int(s * 0.10)),
             (ax + int(s * 0.09), ay1 - int(s * 0.10)),
             (ax, ay1 + int(s * 0.02))]
    d.polygon(arrow, fill=(46, 170, 90, 255))
    return img


def main():
    sizes = [16, 32, 48, 64, 128, 256]
    imgs = [draw(sz) for sz in sizes]
    imgs[-1].save(OUT, format="ICO", sizes=[(sz, sz) for sz in sizes])
    print("icon written:", OUT)


if __name__ == "__main__":
    main()