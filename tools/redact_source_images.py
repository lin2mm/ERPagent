#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把原文档里的架构图做「去公司名」处理后落进仓库 assets/。

用法：
    python tools/redact_source_images.py \
        --src ~/source-docs/docx-media --out assets/case-b

处理规则（用户口径：只去掉公司名 / Logo，界面与内容全部保留）：
  * 标题带里的企业名 —— 用同色同字号的通用名替换（不是打码，版面看不出动过）
  * 其余内容（流程、角色、菜单、权限矩阵、说明文字）原样保留
  * 顺带裁掉四周多余白边，让图在 PPT 里能占满

脚本只做几何与像素处理，不联网；输出为 ASCII 文件名。
"""
import argparse
import os
import shutil

from PIL import Image, ImageDraw, ImageFont

RED = "/tmp/fonts/NotoSansSC-Regular.otf"
BOLD = "/tmp/fonts/NotoSansSC-Bold.otf"

# 文件名 → 说明（用于核对）
MAP = [
    ("image1.png", "arch-01-overview.png", "总体架构（标题带含企业名，需替换）"),
    ("image2.png", "arch-02-one-code.png", "一个码串起一樘门的一生"),
    ("image3.png", "arch-03-permissions.png", "角色 × 菜单权限矩阵"),
    ("image4.png", "arch-04-two-stage-split.png", "两段拆单"),
    ("image5.png", "arch-05-eight-weeks.png", "八周冲刺计划"),
]

# 原始标题 → 替换标题（去掉「盼安门业」这类企业名）
TITLE_SWAP = {
    "arch-01-overview.png": ("盼安门业智能制造一体化系统 · 总体架构",
                             "非标门智能制造一体化系统 · 总体架构"),
}


def dark_bbox(im, y0, y1, thr=420):
    """在 y0..y1 行内找出深色像素的外接框。"""
    px = im.load()
    xs, ys = [], []
    for y in range(y0, min(y1, im.height)):
        for x in range(im.width):
            r, g, b = px[x, y]
            if r + g + b < thr:
                xs.append(x)
                ys.append(y)
    if not xs:
        return None
    return min(xs), min(ys), max(xs), max(ys)


def fit_size(text, target_h, font_path, lo=12, hi=60):
    """找出让文字墨迹高度≈target_h 的字号。"""
    best, best_err = lo, 1e9
    for s in range(lo, hi + 1):
        f = ImageFont.truetype(font_path, s)
        box = f.getbbox(text)
        h = box[3] - box[1]
        err = abs(h - target_h)
        if err < best_err:
            best, best_err = s, err
    return best


def swap_title(im, src_name, out_name):
    """把标题行的企业名换成通用名字，保持居中、字号、颜色一致。"""
    swapped = TITLE_SWAP.get(out_name)
    if not swapped:
        return None
    old, new = swapped
    bb = dark_bbox(im, 0, 80)
    if bb is None:
        raise RuntimeError(f"{src_name}: 找不到标题文字")
    x0, y0, x1, y1 = bb
    th = y1 - y0 + 1
    size = fit_size(new, th, RED)
    f = ImageFont.truetype(RED, size)
    nb = f.getbbox(new)
    nw, nh = nb[2] - nb[0], nb[3] - nb[1]
    # 采样原文字色（取深色像素中出现最多的颜色）
    px = im.load()
    cols = {}
    for y in range(y0, y1 + 1):
        for x in range(x0, x1 + 1):
            c = px[x, y]
            if sum(c) < 420:
                cols[c] = cols.get(c, 0) + 1
    color = max(cols.items(), key=lambda kv: kv[1])[0]
    d = ImageDraw.Draw(im)
    pad = 10
    d.rectangle([x0 - pad, y0 - pad, x1 + pad, y1 + pad], fill=(255, 255, 255))
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    xy = (cx - nw / 2 - nb[0], cy - nh / 2 - nb[1])
    d.text(xy, new, font=f, fill=color)
    return (f"标题替换：{old} → {new}"
            f"（字号 {size}px，色 #{color[0]:02X}{color[1]:02X}{color[2]:02X}，居中）")


def trim(im, keep=14, bg=250):
    """裁掉四周的纯白边（保留 keep px 余量）。"""
    px = im.convert("RGB").load()
    W, H = im.size
    xs, ys = [], []
    for y in range(H):
        for x in range(W):
            r, g, b = px[x, y]
            if not (r >= bg and g >= bg and b >= bg):
                xs.append(x)
                ys.append(y)
    if not xs:
        return im
    x0 = max(0, min(xs) - keep)
    y0 = max(0, min(ys) - keep)
    x1 = min(W, max(xs) + keep + 1)
    y1 = min(H, max(ys) + keep + 1)
    return im.crop((x0, y0, x1, y1))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=os.path.expanduser("~/source-docs/docx-media"))
    ap.add_argument("--out", default="assets/case-b")
    ap.add_argument("--font", default="", help="fontconfig 提供字体时可用；默认 Noto webfont")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)
    for src, dst, note in MAP:
        sp = os.path.join(args.src, src)
        im = Image.open(sp).convert("RGB")
        msg = swap_title(im, src, dst)
        im = trim(im)
        dp = os.path.join(args.out, dst)
        im.save(dp, optimize=True)
        print(f"{dst:<32} {im.width}x{im.height}  {note}"
              + (f"\n    {msg}" if msg else ""))
        assert all(ord(c) < 128 for c in dst)
    print("完成 →", args.out)


if __name__ == "__main__":
    main()
