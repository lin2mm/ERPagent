#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成本地下载页 site/index.html —— 浏览器里直接点，直接下载，不用去 GitHub。

配合 tools/serve_downloads.py 使用：
    python tools/serve_downloads.py          # 起一个静态站点，默认 8000 端口
然后打开预览地址即可，页面上每个文件都是「点一下就开始下载」。

页面由本脚本扫描仓库实时生成：文件大小、是否存在都按当前磁盘内容算，
所以名单/PPT/PDF 更新后重跑一次即可。
"""
import os
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
OUT = os.path.join(ROOT, "site", "index.html")

GROUPS = [
    ("① 客户名单（先看这个）", [
        ("目标企业名单 · Excel", "prospects/Target-Companies.xlsx",
         "28 家初筛名单；三张表：名单 / 评分口径 / 使用说明与获客渠道；表头冻结 + 筛选已开"),
        ("目标企业名单 · CSV", "prospects/Target-Companies.csv", "同一份数据的通用格式"),
        ("目标企业名单 · 详解文档", "docs/08-Prospect-Shortlist.md",
         "评分速览 + 逐家成功面 / 失败面 / 切入点 + 获客渠道"),
    ]),
    ("② 老板版方案（说服用，含系统实景图）", [
        ("PITCH-02 非标定制 · 老板版 PDF", "pdf/PITCH-02-Custom-Door-Boss-Briefing.pdf",
         "19 页，含原文档 5 张架构图（已去企业名）"),
        ("PITCH-02 非标定制 · 可编辑 PPT", "ppt/PITCH-02-Custom-Door-Boss-Briefing.pptx",
         "Keynote 可直接改"),
        ("PITCH-01 卷材 MES · 老板版 PDF", "pdf/PITCH-01-Roll-Material-MES-Boss-Briefing.pdf",
         "16 页（界面截图待补）"),
        ("PITCH-01 卷材 MES · 可编辑 PPT", "ppt/PITCH-01-Roll-Material-MES-Boss-Briefing.pptx", ""),
    ]),
    ("③ 20 分钟面谈版（现场用）", [
        ("PITCH-01 卷材 MES · 8 页 PDF", "pdf/PITCH-01-Short-Roll-Material-MES.pdf", "不带封面，翻开就进正题"),
        ("PITCH-02 非标定制 · 8 页 PDF", "pdf/PITCH-02-Short-Custom-Door.pdf", ""),
    ]),
    ("③b 半天走车间 · 访谈提纲（长三角 S/A 三家）", [
        ("① 江苏赛康医疗设备（张家港）", "interviews/01-SAICOM-Medical-Half-Day-Walkthrough.md",
         "案例 B · S 级 27 分 · 医疗设备与护理家具"),
        ("② 江苏丽岛新材料（常州）", "interviews/02-Lidao-New-Material-Half-Day-Walkthrough.md",
         "案例 A · A 级 25 分 · 彩涂铝卷辊涂，与案例 A 同工艺"),
        ("③ 上海雅轩办公家具（嘉定）", "interviews/03-SpaceIn-Yaxuan-Half-Day-Walkthrough.md",
         "案例 B · S 级 28 分 · 系统家具与工程交付，同城"),
    ]),
    ("④ 客户筛选与打分", [
        ("筛选清单 · A4 打印版 PDF", "pdf/CHECKLIST-Customer-Screening.pdf", "两页，跑车间时拿在手上用"),
        ("筛选清单 · 可编辑模板", "templates/07-Customer-Qualification-Checklist.md", "含设计理由附录"),
    ]),
    ("⑤ 机制版方案（技术底稿，18 页 ×2）", [
        ("PPT-01 卷材 MES · PDF", "pdf/PPT-01.pdf", ""),
        ("PPT-02 非标定制 · PDF", "pdf/PPT-02.pdf", ""),
        ("PPT-01 · 可编辑 PPT", "ppt/PPT-01-Roll-Material-MES-System-Proposal.pptx", ""),
        ("PPT-02 · 可编辑 PPT", "ppt/PPT-02-Custom-Door-Smart-Manufacturing-Plan.pptx", ""),
    ]),
    ("⑥ 说明与评分报告", [
        ("怎么打开、怎么下载", "download/HOW-TO-DOWNLOAD.md", "含直链与常见问题"),
        ("原方案评分报告", "review/01-Source-Materials-Score.md", ""),
        ("成品 PPT 评分报告", "review/02-PPT-Score.md", ""),
        ("本轮交付自评", "review/03-Self-Score-Delivery.md", "下载页 88 / 名单核实 85"),
    ]),
    ("⑦ 一次全下", [
        ("全部材料", "download/ALL-MATERIALS.zip", "PPT + PDF + Excel + 逐页 PNG + 文档"),
    ]),
]

PREVIEWS = [
    ("PITCH-01 老板版预览总图", "ppt/preview/PITCH-01-Roll-Material-MES-Preview.png"),
    ("PITCH-02 老板版预览总图（19 页）", "ppt/preview/PITCH-02-Custom-Door-Preview.png"),
    ("PITCH-01 面谈版预览总图", "ppt/preview/PITCH-01-Short-Preview-8pages.png"),
    ("PITCH-02 面谈版预览总图", "ppt/preview/PITCH-02-Short-Preview-8pages.png"),
    ("PPT-01 机制版预览总图", "ppt/preview/PPT-01-Preview-18pages.png"),
    ("PPT-02 机制版预览总图", "ppt/preview/PPT-02-Preview-18pages.png"),
]

CSS = """
:root { --ink:#0F1B2D; --body:#33415C; --muted:#7A8699; --line:#DDE3EC; --soft:#F5F7FB;
        --acc:#2C5AA0; --ok:#1E6F50; }
* { box-sizing: border-box; }
body { margin:0; background:#EEF1F6; color:var(--body);
       font: 15px/1.7 -apple-system, "PingFang SC", "Microsoft YaHei", "Helvetica Neue", Arial, sans-serif; }
.wrap { max-width: 1040px; margin: 0 auto; padding: 34px 20px 70px; }
h1 { color:var(--ink); font-size:27px; margin:0 0 6px; letter-spacing:.3px; }
.sub { color:var(--muted); margin:0 0 26px; font-size:13.5px; }
h2 { color:var(--ink); font-size:17px; margin:30px 0 12px; padding-bottom:8px; border-bottom:2px solid var(--line); }
.grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(310px,1fr)); gap:12px; }
a.card { display:block; text-decoration:none; background:#fff; border:1px solid var(--line);
         border-radius:11px; padding:14px 16px; transition:.15s; }
a.card:hover { border-color:var(--acc); box-shadow:0 6px 18px rgba(20,40,80,.10); transform:translateY(-1px); }
a.card .t { color:var(--ink); font-weight:700; font-size:14.5px; display:flex; justify-content:space-between; gap:10px; }
a.card .d { color:var(--muted); font-size:12.5px; margin-top:5px; }
a.card .f { color:#9AA6B6; font-size:11.5px; margin-top:7px; word-break:break-all; }
.pill { background:var(--soft); color:var(--body); border-radius:20px; padding:2px 9px; font-size:11.5px;
        font-weight:600; white-space:nowrap; }
.pill.zip { background:#E7F1EC; color:var(--ok); }
.note { background:#fff; border:1px solid var(--line); border-left:4px solid var(--acc);
        border-radius:10px; padding:14px 18px; margin:26px 0 0; font-size:13.5px; }
.note b { color:var(--ink); }
.imgs { display:grid; grid-template-columns:repeat(auto-fill, minmax(240px,1fr)); gap:12px; }
.imgs a { display:block; background:#fff; border:1px solid var(--line); border-radius:10px; padding:10px;
          text-decoration:none; }
.imgs img { width:100%; border-radius:6px; border:1px solid var(--line); }
.imgs .t { color:var(--ink); font-size:12.8px; font-weight:600; margin-top:8px; }
footer { color:var(--muted); font-size:12px; margin-top:34px; }
"""


ZIP_PATH = os.path.join(ROOT, "download", "ALL-MATERIALS.zip")


def zip_count():
    try:
        return len(zipfile.ZipFile(ZIP_PATH).namelist())
    except Exception:
        return 0


def size_str(p):
    if not os.path.exists(p):
        return "文件缺失"
    n = os.path.getsize(p)
    return f"{n/1024/1024:.1f} MB" if n >= 1024 * 1024 else f"{n/1024:.0f} KB"


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def card(title, path, desc):
    if path.endswith("ALL-MATERIALS.zip"):
        title = f"{title}（{zip_count()} 个文件）"
    exists = os.path.exists(os.path.join(ROOT, path))
    cls = "pill zip" if path.endswith(".zip") else "pill"
    d = f'<div class="d">{esc(desc)}</div>' if desc else ""
    return (f'<a class="card" href="/{path}" download>'
            f'<div class="t"><span>{esc(title)}</span><span class="{cls}">{size_str(os.path.join(ROOT, path))}</span></div>'
            f'{d}<div class="f">{esc(path)}</div></a>')


def build():
    parts = []
    for gname, items in GROUPS:
        cards = "\n".join(card(t, p, d) for t, p, d in items)
        parts.append(f'<h2>{esc(gname)}</h2>\n<div class="grid">\n{cards}\n</div>')
    imgs = "\n".join(
        f'<a href="/{p}" download><img src="/{p}" alt="{esc(t)}"><div class="t">{esc(t)}</div></a>'
        for t, p in PREVIEWS if os.path.exists(os.path.join(ROOT, p)))
    html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>交付物下载页 · 智能制造方案</title><style>{CSS}</style></head>
<body><div class="wrap">
<h1>交付物下载页</h1>
<p class="sub">每个文件点一下就开始下载（<b>不需要 GitHub、不需要登录</b>）。
本页由 <code>tools/build_download_page.py</code> 扫描仓库生成，{sum(len(i) for _, i in GROUPS)} 个文件。</p>
{chr(10).join(parts)}
<h2>⑧ 预览总图（先看这个决定要不要下 PDF）</h2>
<div class="imgs">
{imgs}
</div>
<div class="note">
<b>给同事的用法：</b>需要哪份点哪份；不确定就先下
<a href="/download/ALL-MATERIALS.zip" download>ALL-MATERIALS.zip</a>（{size_str(os.path.join(ROOT, 'download/ALL-MATERIALS.zip'))}，一次拿全）。
<br><b>注意：</b>本页只在会话进行时有效（本地预览服务）；长期保存请下载到本机或走 GitHub 直链。
</div>
<footer>生成时间戳由文件内容决定 · 路径均为仓库相对路径 · 全部 ASCII 文件名，避免下载后乱码</footer>
</div></body></html>"""
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    return OUT


if __name__ == "__main__":
    p = build()
    print("写入", p, f"{os.path.getsize(p)/1024:.1f} KB")
