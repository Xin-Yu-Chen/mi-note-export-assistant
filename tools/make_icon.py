#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成应用图标 assets/icon.ico（无需联网，纯 Pillow 绘制）。运行：py tools/make_icon.py"""
import os
from PIL import Image, ImageDraw

SIZE = 256
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "assets", "icon.ico")

TOP = (58, 134, 255)      # 亮蓝
BOTTOM = (0, 180, 216)    # 青
LINE = (156, 195, 232)    # 文档内浅蓝线
ARROW = (27, 111, 224)    # 箭头蓝


def vgradient(size, top, bottom):
    img = Image.new("RGB", (size, size), top)
    px = img.load()
    for y in range(size):
        t = y / (size - 1)
        row = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
        for x in range(size):
            px[x, y] = row
    return img


def rounded_mask(size, radius):
    m = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return m


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    bg = vgradient(SIZE, TOP, BOTTOM)
    icon = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    icon.paste(bg, (0, 0), rounded_mask(SIZE, 56))
    d = ImageDraw.Draw(icon)

    # 白色文档卡片
    d.rounded_rectangle([72, 46, 184, 172], radius=14, fill=(255, 255, 255, 255))
    # 文档里的三行文字线
    for i, y in enumerate((76, 100, 124)):
        x2 = 168 if i < 2 else 140
        d.rounded_rectangle([88, y, x2, y + 9], radius=4, fill=LINE)

    # 导出圆（叠在文档下沿）
    d.ellipse([92, 150, 164, 222], fill=(255, 255, 255, 255), outline=ARROW, width=6)
    # 向下箭头（竖条 + 三角）
    d.rounded_rectangle([124, 166, 132, 190], radius=3, fill=ARROW)
    d.polygon([(128, 204), (112, 186), (144, 186)], fill=ARROW)

    icon.save(OUT, sizes=[(16, 16), (32, 32), (48, 48), (64, 64),
                          (128, 128), (256, 256)])
    # 同时存一张 256 png 便于 README 预览
    png = OUT.replace(".ico", ".png")
    icon.save(png)
    print("图标已生成:", OUT)


if __name__ == "__main__":
    main()
