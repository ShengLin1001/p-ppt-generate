# -*- coding: utf-8 -*-
"""插图入页前裁掉四周白边，另存为 <原名>-trim.png。

论文图自带的白边会在面板里叠成大片空白（面板看起来"空"的主因之一），
而面板尺寸是按图的宽高比定的，白边不裁，宽高比就是错的。
"""

import argparse
import sys
from pathlib import Path

from PIL import Image, ImageChops

SUFFIX_TRIM = "-trim"
# 近白阈值：JPEG/抗锯齿边缘不是纯白，差值 ≤ 此值都当背景
THRESHOLD_WHITE = 12


def fail(msg):
    print("❌ ERROR: " + msg)
    raise SystemExit(1)


def get_lpath_image(path_input):
    if path_input.is_file():
        return [path_input]
    if not path_input.is_dir():
        fail("输入不存在 " + str(path_input))
    return sorted(p for p in path_input.iterdir()
                  if p.suffix.lower() in (".png", ".jpg", ".jpeg")
                  and not p.stem.endswith(SUFFIX_TRIM))


def get_bbox_content(image):
    """非白内容的包围盒；整张是白图时返回 None。"""
    rgb = image.convert("RGB")
    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255)))
    mask = diff.convert("L").point(lambda v: 255 if v > THRESHOLD_WHITE else 0)
    return mask.getbbox()


def trim_image(path_image, pad):
    """裁边并另存，返回 (新路径, 原尺寸, 新尺寸)；无内容返回 None。"""
    image = Image.open(path_image)
    bbox = get_bbox_content(image)
    if bbox is None:
        return None
    left, top, right, bottom = bbox
    box = (max(0, left - pad), max(0, top - pad),
           min(image.width, right + pad), min(image.height, bottom + pad))
    path_out = path_image.with_name(path_image.stem + SUFFIX_TRIM + ".png")
    image.crop(box).save(path_out)
    return path_out, image.size, (box[2] - box[0], box[3] - box[1])


def build_parser():
    parser = argparse.ArgumentParser(description="裁掉插图四周白边")
    parser.add_argument("-input", required=True, help="单张图或图片目录")
    parser.add_argument("-pad", type=int, default=8, help="裁后保留的边距像素")
    return parser


args = build_parser().parse_args()
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

path_input = Path(args.input).resolve()
lpath_image = get_lpath_image(path_input)
print("📁 %s ｜ %d 张" % (path_input, len(lpath_image)))

lempty = []
for path_image in lpath_image:
    result = trim_image(path_image, args.pad)
    if result is None:
        lempty.append(path_image.name)
        print("  ⚠️  %s 全白，跳过" % path_image.name)
        continue
    path_out, size_old, size_new = result
    ratio = size_new[0] * size_new[1] / float(size_old[0] * size_old[1])
    print("  ✅ %s  %dx%d → %dx%d（保留 %.0f%%）"
          % (path_out.name, size_old[0], size_old[1], size_new[0], size_new[1], ratio * 100))

print("📊 %d 张 ｜ 裁边 %d ｜ 全白 %d" % (len(lpath_image), len(lpath_image) - len(lempty), len(lempty)))
print("🎉 done")
