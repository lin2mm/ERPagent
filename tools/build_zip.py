#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把要交付给客户的成品打成一个 zip（download/ALL-MATERIALS.zip）。

用法：
    python tools/build_zip.py

只打「成品」，不打源码工具与 git 相关文件：
    README.md · docs/ · templates/ · spec/ · review/ · prospects/(csv+xlsx)
    ppt/*.pptx · ppt/preview/*.png · pdf/*.pdf
    download/<各 PNG 目录>/ · download/HOW-TO-DOWNLOAD.md
    assets/(已脱敏插图)
zip 内不含 zip 自身。
"""
import os
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
OUT = os.path.join(ROOT, "download", "ALL-MATERIALS.zip")
SKIP_EXT = {".zip", ".pyc"}
SKIP_DIRS = {"__pycache__", ".git"}


def collect():
    files = ["README.md"]
    for d in ("docs", "templates", "spec", "review", "prospects", "assets"):
        base = os.path.join(ROOT, d)
        for dirpath, dirnames, filenames in os.walk(base):
            dirnames[:] = [x for x in dirnames if x not in SKIP_DIRS]
            for fn in sorted(filenames):
                if os.path.splitext(fn)[1] in SKIP_EXT:
                    continue
                files.append(os.path.relpath(os.path.join(dirpath, fn), ROOT))
    for d in ("ppt", "pdf"):
        base = os.path.join(ROOT, d)
        for dirpath, dirnames, filenames in os.walk(base):
            for fn in sorted(filenames):
                if os.path.splitext(fn)[1] in SKIP_EXT:
                    continue
                files.append(os.path.relpath(os.path.join(dirpath, fn), ROOT))
    dow = os.path.join(ROOT, "download")
    for dirpath, dirnames, filenames in os.walk(dow):
        for fn in sorted(filenames):
            p = os.path.join(dirpath, fn)
            if p == OUT or os.path.splitext(fn)[1] in SKIP_EXT:
                continue
            files.append(os.path.relpath(p, ROOT))
    return sorted(set(files))


def main():
    files = collect()
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for f in files:
            z.write(os.path.join(ROOT, f), f)
    with zipfile.ZipFile(OUT) as z:
        bad = z.testzip()
        n = len(z.namelist())
        assert bad is None, bad
        assert n == len(files), (n, len(files))
    size = os.path.getsize(OUT) / 1024 / 1024
    print(f"ALL-MATERIALS.zip: {len(files)} 个文件 · {size:.1f} MB")
    for f in files:
        assert all(ord(c) < 128 for c in f), f
    print("文件名全部 ASCII ✓")


if __name__ == "__main__":
    main()
