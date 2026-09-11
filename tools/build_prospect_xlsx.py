#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""由 prospects/Target-Companies.csv 生成可下载的 Excel 名单。

用法：
    python tools/build_prospect_xlsx.py

产出：prospects/Target-Companies.xlsx（三张表：名单 / 评分口径 / 使用说明）
CSV 是唯一数据源，改名单改 CSV 再重跑本脚本即可。

为什么给 Excel 而不是 CSV：
  · 可筛选、可排序、可冻结表头、可直接打印
  · 销售人员可以在里面加自己的备注列，不用装任何软件
"""
import csv
import os

import xlsxwriter

ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "..")
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
CSV = os.path.join(ROOT, "prospects", "Target-Companies.csv")
OUT = os.path.join(ROOT, "prospects", "Target-Companies.xlsx")

INK = "#0F1B2D"
BODY = "#33415C"
MUTED = "#7A8699"
LINE = "#D8DEE9"
SOFT = "#F2F5F9"

TIER_STYLE = {
    "S": ("#1E6F50", "#FFFFFF"),
    "A": ("#2C5AA0", "#FFFFFF"),
    "B": ("#5A6472", "#FFFFFF"),
    "C": ("#9AA3B0", "#FFFFFF"),
}

COLS = [("优先级", 8), ("企业名称", 30), ("省市", 8), ("城市", 13), ("网址", 26),
        ("行业细分", 12), ("主营业务", 34), ("适配案例", 10),
        ("初筛分\n(企业侧/35)", 11), ("匹配点（成功面）", 52),
        ("风险点（失败面 · 一律待核实）", 52), ("建议切入点", 38),
        ("接触路径", 20), ("信息可靠度", 10)]


def build():
    with open(CSV, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    head, data = rows[0], rows[1:]

    wb = xlsxwriter.Workbook(OUT, {"default_date_format": "yyyy-mm-dd"})
    fmt = dict(
        title=wb.add_format({"font_size": 15, "bold": True, "font_color": INK}),
        sub=wb.add_format({"font_size": 10, "font_color": MUTED}),
        head=wb.add_format({"bold": True, "font_size": 10, "font_color": "#FFFFFF",
                            "bg_color": INK, "align": "center", "valign": "vcenter",
                            "text_wrap": True, "border": 1, "border_color": INK}),
        cell=wb.add_format({"font_size": 10, "font_color": BODY, "valign": "top",
                            "text_wrap": True, "border": 1, "border_color": LINE}),
        num=wb.add_format({"font_size": 10, "font_color": BODY, "valign": "top",
                           "align": "center", "border": 1, "border_color": LINE}),
        link=wb.add_format({"font_size": 9, "font_color": "#2C5AA0", "valign": "top",
                            "text_wrap": True, "border": 1, "border_color": LINE}),
        band=wb.add_format({"bold": True, "font_size": 10.5, "font_color": INK,
                            "bg_color": SOFT, "border": 1, "border_color": LINE}),
        sect=wb.add_format({"bold": True, "font_size": 12, "font_color": INK}),
        note=wb.add_format({"font_size": 10, "font_color": BODY, "text_wrap": True,
                            "valign": "top"}),
    )

    # ── 表 1：名单 ────────────────────────────────────────────────
    ws = wb.add_worksheet("名单")
    ws.set_column(0, 0, COLS[0][1])
    for i, (_, w) in enumerate(COLS):
        ws.set_column(i, i, w)
    ws.write(0, 0, "目标企业名单（初筛）· 按筛选标准打分排序 · 公开信息，非尽调结论",
             fmt["title"])
    ws.write(1, 0, "分数只用于「先谈谁、后谈谁」；规模看员工数（80–800 人），不按营收判断。"
                   "老板侧 B1–B7 必须面谈才能打分，见「评分口径」表。", fmt["sub"])
    ws.set_row(0, 22)
    ws.set_row(1, 16)
    hr = 2
    for c, (name, _) in enumerate(COLS):
        ws.write(hr, c, name, fmt["head"])
    ws.set_row(hr, 30)

    for r, row in enumerate(data, start=hr + 1):
        tier = row[0]
        bg, fg = TIER_STYLE.get(tier, (SOFT, INK))
        tfmt = wb.add_format({"bold": True, "font_size": 11, "font_color": fg,
                              "bg_color": bg, "align": "center", "valign": "vcenter",
                              "border": 1, "border_color": LINE})
        for c, val in enumerate(row):
            f = fmt["num"] if c in (0, 2, 3, 7, 8, 13) else fmt["cell"]
            if c == 0:
                f = tfmt
            elif c in (4, 5) and val:
                f = fmt["link"]
            ws.write(r, c, val, f)
        ws.set_row(r, 62)

    ws.freeze_panes(hr + 1, 2)
    ws.autofilter(hr, 0, hr + len(data), len(COLS) - 1)
    ws.set_landscape()
    ws.set_paper(9)                      # A4
    ws.fit_to_pages(1, 0)
    ws.set_margins(0.3, 0.3, 0.4, 0.4)
    ws.repeat_rows(hr)

    # ── 表 2：评分口径 ────────────────────────────────────────────
    ws2 = wb.add_worksheet("评分口径")
    ws2.set_column(0, 0, 26)
    ws2.set_column(1, 1, 10)
    ws2.set_column(2, 2, 86)
    r = 0
    ws2.write(r, 0, "打分口径（与 templates/07 客户筛选清单一致）", fmt["title"])
    r += 2
    blocks = [
        ("企业侧 · 公开信息就能打（满分 35）", [
            ("E1 规模匹配", "5", "1–10 亿营收 或 80–800 人：正中甜点区（材料/工程类营收被原材料放大，优先看人数）"),
            ("E2 痛点具体性", "5", "能说出具体事件与金额：被罚过 / 压了几百万 / 赔过 / 丢过标"),
            ("E3 外部压力", "5", "已有硬要求：大客户溯源、招标资质、合规检查、资本化"),
            ("E4 数据基础", "5", "有物料编码，且有打标 / 条码设备 —— 说明「码」已经存在，只差把它激活"),
            ("E5 业务稳定性", "5", "订单饱满、正在扩产（没空折腾但有增长压力 = 最好的时机）"),
            ("E6 付费能力", "5", "有预算科目、接受按里程碑付款"),
            ("E7 行业价值密度", "5", "非标定制 / 原材料占比高 / 有追溯要求（本名单据此选行业）"),
        ]),
        ("老板侧 · 必须面谈（满分 35，不在这张表里）", [
            ("B1 亲自下场意愿 ★", "5", "最强单一预测因子：愿意把关键用户按在会议室签字、参加双周演示"),
            ("B2 说得出瓶颈数字", "5", "报得出具体数字，且知道卡在哪道工序"),
            ("B3 数字化体感", "5", "手机上看数据，主动问「能不能手机上批」"),
            ("B4 吃过亏且记得住", "5", "能讲出一次具体的损失事故"),
            ("B5 扩张野心", "5", "要接大客户 / 扩产 / 做品牌，才愿意为「管理能力」付费"),
            ("B6 年龄与精力", "5", "35–55 岁、仍在一线管事"),
            ("B7 决策效率", "5", "当场能定方向、指定对接人"),
        ]),
        ("定级与投入", [
            ("S 级 ≥ 56", "", "立即投入：同行业活案例参观 → 4 周内做出只读驾驶舱 → 谈分期与里程碑"),
            ("A 级 42–55", "", "重点培育：先做半天走车间 + 1 页症断书，建立信任后再谈主线"),
            ("B 级 28–41", "", "保持观察：定期发行业案例、邀请参加活动，等外部压力出现"),
            ("C 级 < 28", "", "礼貌退出；命中一票否决或 ★ 的，同样按 C 级处理"),
        ]),
        ("两条修正规则", [
            ("规则一", "", "总分高但老板不愿亲自下场 → 按 B 级处理（别在没人的项目上耗时间）"),
            ("规则二", "", "总分中等但老板肯下场、且有外部压力 → 可按 A 级投入"),
        ]),
        ("这张表里的分数是什么", [
            ("初筛分", "", "本名单只给了「企业侧」的公开信息估算分（18–29 / 35），用来排序与分配拜访精力"),
            ("为什么不满分", "", "公开信息拿不到 E2/E6 的真实答案；老板侧 35 分完全空白 —— 见面才能补"),
            ("怎么用", "", "先用它挑出 5–8 家，拿出半天走一次车间，把 E 与 B 补齐，再决定投入多少"),
        ]),
    ]
    for title, items in blocks:
        ws2.write(r, 0, title, fmt["sect"])
        r += 1
        for a, b, c in items:
            ws2.write(r, 0, a, fmt["band"])
            ws2.write(r, 1, b, fmt["band"])
            ws2.write(r, 2, c, fmt["note"])
            ws2.set_row(r, 30)
            r += 1
        r += 1

    # ── 表 3：使用说明与获客渠道 ──────────────────────────────────
    ws3 = wb.add_worksheet("使用说明与获客渠道")
    ws3.set_column(0, 0, 22)
    ws3.set_column(1, 1, 96)
    r = 0
    ws3.write(r, 0, "怎么用这份名单（三步）", fmt["title"])
    r += 2
    guide = [
        ("第 1 步 · 10 分钟桌面核实", "打开官网/工商信息，核对三件事：员工规模、主营产品、有没有打标/条码设备。"
                                  "任何一条对不上，先把这家降到 B 级观察。"),
        ("第 2 步 · 3 个探路电话", "不推销，只问三句：最近订单排到什么时候？有没有因为交期或批次被客户罚过/退过？"
                              "老板自己管不管生产？三句里有两句是具体事件 → 直接约半天走车间。"),
        ("第 3 步 · 五步阶梯", "半天走车间 → 1 周症断书 → 2 周只读驾驶舱 → 6–8 周试点 → 推广。"
                           "每一步都用上一份成品：16 页老板版讲价值，8 页面谈版现场讲，筛选清单当场打分。"),
    ]
    for a, b in guide:
        ws3.write(r, 0, a, fmt["band"])
        ws3.write(r, 1, b, fmt["note"])
        ws3.set_row(r, 44)
        r += 1
    r += 1
    ws3.write(r, 0, "四种持续获客渠道（名单会过时，渠道不会）", fmt["sect"])
    r += 1
    chans = [
        ("招标中标名单", "中国政府采购网 ccgp.gov.cn / 各省招标公告平台，按行业搜「家具」「防火门」「彩涂板」。"
                    "中标公告里同时有甲方、乙方、金额、工期 —— 工期违约条款抄下来，就是方案第一页的「账」。"),
        ("产业带集群", "博兴涂镀（彩涂板）、永康/任丘（门业）、佛山南海（门窗家具）、奉贤（展柜）、阜城（集成房屋）。"
                  "一个镇一条产业链，做成一家就能顺着熟人网络复制。"),
        ("行业展会", "CHCC 全国医院建设大会、FBC 门窗幕墙、广交会、商业空间展。"
                 "展会抓的是「正在花钱做市场」的企业，通常也正在扩产。"),
        ("招聘信息", "搜「MES 实施」「信息化主管」「数字化专员」。在招这类岗位 = 已有预算科目，"
                 "优先级直接从 B 提到 A，且对接人就是未来的项目负责人。"),
    ]
    for a, b in chans:
        ws3.write(r, 0, a, fmt["band"])
        ws3.write(r, 1, b, fmt["note"])
        ws3.set_row(r, 60)
        r += 1
    r += 1
    ws3.write(r, 0, "免责与保密", fmt["sect"])
    r += 1
    ws3.write(r, 0, "说明", fmt["band"])
    ws3.write(r, 1, "本名单全部来自公开信息（公司官网、上市公司公告、权威媒体报道、招标公告、"
                    "行业榜单与工商公开信息），只做初筛与排序，不是尽调结论；"
                    "所有「风险点」都是待核实项，不构成对任何企业的负面判断。"
                    "本文件所在仓库为公开仓库，如用于商业开发请先移入私有仓库。", fmt["note"])
    ws3.set_row(r, 66)

    wb.close()
    return OUT


if __name__ == "__main__":
    p = build()
    print("写入", p, f"{os.path.getsize(p)/1024:.1f} KB")
