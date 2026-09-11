#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""由 prospects/Target-Companies.csv 生成可下载的 Excel 名单（v2）。

    python tools/build_prospect_xlsx.py

产出 prospects/Target-Companies.xlsx，四张表：
    名单              —— 28 家 × 22 列，含官网核实、联系方式、触达路径
    本轮核实纪要      —— 核实方法、结果统计、逐家修正清单
    评分口径          —— E1–E7 / B1–B7 / 定级线 / 修正规则
    使用说明与获客渠道

CSV 是唯一数据源：改名单改 CSV，重跑本脚本即可。
"""
import csv
import os
import re

import xlsxwriter

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
CSV = os.path.join(ROOT, "prospects", "Target-Companies.csv")
OUT = os.path.join(ROOT, "prospects", "Target-Companies.xlsx")

INK, BODY, MUTED, LINE, SOFT = "#0F1B2D", "#33415C", "#7A8699", "#D8DEE9", "#F2F5F9"
TIER_STYLE = {"S": "#1E6F50", "A": "#2C5AA0", "B": "#5A6472", "C": "#9AA3B0"}

WIDTHS = [7, 30, 30, 8, 12, 26, 40, 44, 24, 20, 30, 44, 11, 26, 9, 9, 40, 40, 30, 30, 9, 11]
WRAP_COLS = {1, 2, 5, 6, 7, 8, 9, 10, 11, 13, 16, 17, 18, 19}
CENTER_COLS = {0, 3, 4, 12, 14, 15, 20, 21}


def fmt(wb):
    return dict(
        head=wb.add_format({"bold": True, "font_size": 10, "font_color": "#FFFFFF",
                            "bg_color": INK, "align": "center", "valign": "vcenter",
                            "text_wrap": True, "border": 1, "border_color": INK}),
        cell=wb.add_format({"font_size": 9.5, "font_color": BODY, "valign": "top",
                            "text_wrap": True, "border": 1, "border_color": LINE}),
        num=wb.add_format({"font_size": 9.5, "font_color": BODY, "valign": "top",
                           "align": "center", "border": 1, "border_color": LINE}),
        link=wb.add_format({"font_size": 9, "font_color": "#2C5AA0", "valign": "top",
                            "text_wrap": True, "border": 1, "border_color": LINE}),
        adv=wb.add_format({"font_size": 9.5, "bold": True, "font_color": "#1E5E8A",
                           "valign": "top", "text_wrap": True, "border": 1,
                           "border_color": LINE, "bg_color": "#EAF3FA"}),
        band=wb.add_format({"bold": True, "font_size": 10, "font_color": INK,
                            "bg_color": SOFT, "border": 1, "border_color": LINE,
                            "valign": "vcenter"}),
        sect=wb.add_format({"bold": True, "font_size": 12, "font_color": INK}),
        note=wb.add_format({"font_size": 10, "font_color": BODY, "text_wrap": True,
                            "valign": "top"}),
        title=wb.add_format({"font_size": 14, "bold": True, "font_color": INK}),
    )


def sheet_list(wb, f, rows):
    ws = wb.add_worksheet("名单")
    for i, w in enumerate(WIDTHS):
        ws.set_column(i, i, w)
    head = rows[0]
    for c, name in enumerate(head):
        ws.write(0, c, name, f["head"])
    ws.set_row(0, 34)
    for r, row in enumerate(rows[1:], start=1):
        tier = row[0]
        tfmt = wb.add_format({"bold": True, "font_size": 11, "font_color": "#FFFFFF",
                              "bg_color": TIER_STYLE.get(tier, SOFT), "align": "center",
                              "valign": "vcenter", "border": 1, "border_color": LINE})
        for c, val in enumerate(row):
            if c == 0:
                cell = tfmt
            elif c == 1:
                cell = f["adv"]
            elif c == 5 and val.startswith("http"):
                ws.write_url(r, c, val, f["link"], val)      # 官网可点
                continue
            elif c in CENTER_COLS:
                cell = f["num"]
            else:
                cell = f["cell"]
            if c == 9:                                        # 邮箱可点（整格只放一个邮箱时）
                one = re.fullmatch(r"[\w.+-]+@[\w-]+\.[\w.]+", val.strip())
                if one:
                    ws.write_url(r, c, "mailto:" + val.strip(), f["link"], val.strip())
                    continue
            ws.write(r, c, val, cell)
        ws.set_row(r, 96)
    ws.freeze_panes(1, 3)
    ws.autofilter(0, 0, len(rows) - 1, len(head) - 1)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(0)
    return ws


CORRECTIONS = [
    ("广东华燊实业（VOU+卫域）", "城市从【东莞】修正为【广州花都区】；补三基地地址与三条专线电话"),
    ("浙江金凯德安防科技", "主体不在【永康】而在【武义】；集团 1200+ 人、六大厂区 → 优先级 B → C"),
    ("天津市新宇彩板", "员工 2000+、1100 亩、国家级单项冠军 → B → C（只能做标杆）"),
    ("佛山市南海新达高梵", "员工 1000+、年产值 10 亿 → A → B（只谈车间那一段）"),
    ("广州市万开家具", "无官网、公开信息薄，推广页称 700 人 → A → B"),
    ("任丘市益德门业", "厂址实为【河间市边边工业区】；产品扩为被动门/防火门/医用门三大系列 → B → A"),
    ("浙江香乡门业", "官网：员工 500+、年产 30 万樘防火门 → B → A"),
    ("上海大兆展示", "官网：员工 120 人、7 个项目组 → B → A（正中甜点区）"),
    ("辽宁中安华泰防火门业", "注册资本仅 500 万、无官网 → B → C"),
    ("山东辰贝驰新材料 / 浙江群邦门业", "检索不到企业主体页 → 建议移出（先按工商全称核对）"),
    ("浙江德艺门业", "官网未找到 → 改走永康国际门业博览会现场核实"),
    ("山东好运来集成房屋", "注册资本 300 万偏小、来源为推广文 → 维持 A 但先电话核实厂房与在制项目"),
]


def sheet_verify(wb, f):
    ws = wb.add_worksheet("本轮核实纪要")
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 96)
    ws.write(0, 0, "官网主营反向验证 · 本轮纪要", f["title"])
    ws.write(1, 0, "方法：只看企业官网（首页/关于我们/产品/联系方式）与工商公示信息；"
                   "官网写的口径优先于黄页与软文；核不上的一律写「待核实」，不做断言。", f["note"])
    ws.set_row(1, 30)
    ws.write(3, 0, "核实结果", f["sect"])
    stats = [("覆盖企业", "28 家（全部逐家过一遍）"),
             ("拿到官网链接", "20 家（其余 8 家：无官网或只有 B2B 页面）"),
             ("拿到企业电话", "25 家"), ("拿到企业邮箱", "14 家"),
             ("发现需要修正", "12 条（见下表）"),
             ("结论变化", "4 家升档（益德 / 香乡 / 大兆 + 邑通偏强）、5 家降档、2 家建议移出")]
    r = 4
    for a, b in stats:
        ws.write(r, 0, a, f["band"])
        ws.write(r, 1, b, f["note"])
        ws.set_row(r, 20)
        r += 1
    r += 1
    ws.write(r, 0, "逐家修正清单", f["sect"])
    r += 1
    ws.write(r, 0, "企业", f["band"])
    ws.write(r, 1, "修正内容", f["band"])
    r += 1
    for a, b in CORRECTIONS:
        ws.write(r, 0, a, f["note"])
        ws.write(r, 1, b, f["note"])
        ws.set_row(r, 30)
        r += 1
    r += 1
    ws.write(r, 0, "怎么用这些联系方式", f["sect"])
    r += 1
    ws.write(r, 0, "三条规矩", f["band"])
    ws.write(r, 1, "① 先打电话确认「是不是这家在管生产」，别直接谈软件；② 有政采/投标专线的走专线，"
                   "对口人就是项目负责人；③ 邮箱只发案例 PDF（老板版 16/19 页），不发方案全文。", f["note"])
    ws.set_row(r, 46)


def sheet_rubric(wb, f):
    ws = wb.add_worksheet("评分口径")
    ws.set_column(0, 0, 26)
    ws.set_column(1, 1, 10)
    ws.set_column(2, 2, 86)
    r = 0
    ws.write(r, 0, "打分口径（与 templates/07 客户筛选清单一致）", f["title"])
    r += 2
    blocks = [
        ("企业侧 · 公开信息就能打（满分 35）", [
            ("E1 规模匹配", "5", "1–10 亿营收 或 80–800 人（材料/工程类营收被原材料放大，优先看人数）"),
            ("E2 痛点具体性", "5", "能说出具体事件与金额：被罚过 / 压了几百万 / 赔过 / 丢过标"),
            ("E3 外部压力", "5", "已有硬要求：大客户溯源、招标资质、合规检查、资本化"),
            ("E4 数据基础", "5", "有物料编码，且有打标 / 条码设备"),
            ("E5 业务稳定性", "5", "订单饱满、正在扩产"),
            ("E6 付费能力", "5", "有预算科目、接受按里程碑付款"),
            ("E7 行业价值密度", "5", "非标定制 / 原材料占比高 / 有追溯要求"),
        ]),
        ("老板侧 · 必须面谈（满分 35，不在这张表里）", [
            ("B1 亲自下场意愿 ★", "5", "最强单一预测因子：愿意把关键用户按在会议室签字、参加双周演示"),
            ("B2 说得出瓶颈数字", "5", "报得出具体数字，且知道卡在哪道工序"),
            ("B3 数字化体感", "5", "手机上看数据，主动问「能不能手机上批」"),
            ("B4 吃过亏且记得住", "5", "能讲出一次具体的损失事故"),
            ("B5 扩张野心", "5", "要接大客户 / 扩产 / 做品牌"),
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
            ("规则一", "", "总分高但老板不愿亲自下场 → 按 B 级处理"),
            ("规则二", "", "总分中等但老板肯下场、且有外部压力 → 可按 A 级投入"),
        ]),
        ("名单里的分数是什么", [
            ("初筛分", "", "只给了「企业侧」公开信息估算分（18–29 / 35），用来排序与分配拜访精力"),
            ("为什么不满分", "", "公开信息拿不到 E2/E6 的真实答案；老板侧 35 分完全空白 —— 见面才能补"),
            ("官网核实的作用", "", "先去掉名字与网址对不上的、规模超出的、主体查不到的，再决定投入"),
        ]),
    ]
    for title, items in blocks:
        ws.write(r, 0, title, f["sect"])
        r += 1
        for a, b, c in items:
            ws.write(r, 0, a, f["band"])
            ws.write(r, 1, b, f["band"])
            ws.write(r, 2, c, f["note"])
            ws.set_row(r, 30)
            r += 1
        r += 1
    return ws


def sheet_howto(wb, f):
    ws = wb.add_worksheet("使用说明与获客渠道")
    ws.set_column(0, 0, 22)
    ws.set_column(1, 1, 96)
    r = 0
    ws.write(r, 0, "怎么用这份名单（三步）", f["title"])
    r += 2
    for a, b in [
        ("第 1 步 · 桌面核实（10 分钟/家）", "打开官网，核对员工规模、主营产品、有没有打标/条码设备；"
                                        "对不上就先降到 B 级观察。"),
        ("第 2 步 · 3 个探路电话", "不推销，只问三句：订单排到什么时候？有没有因交期或批次被罚过/退过？老板自己管不管生产？"
                              "三句里两句是具体事件 → 直接约半天走车间。"),
        ("第 3 步 · 五步阶梯", "半天走车间 → 1 周症断书 → 2 周只读驾驶舱 → 6–8 周试点 → 推广。"
                           "每步配一份成品：19 页老板版讲价值，8 页版现场讲，筛选清单当场打分。"),
    ]:
        ws.write(r, 0, a, f["band"])
        ws.write(r, 1, b, f["note"])
        ws.set_row(r, 44)
        r += 1
    r += 1
    ws.write(r, 0, "四种持续获客渠道", f["sect"])
    r += 1
    for a, b in [
        ("招标中标名单", "ccgp.gov.cn 与各省招标平台，搜「家具」「防火门」「彩涂板」。中标公告里有甲方、乙方、金额、工期 —— "
                    "工期违约条款抄下来，就是方案第一页的「账」。"),
        ("产业带集群", "博兴涂镀（彩涂板）、永康/任丘-河间（门业）、佛山南海（门窗家具）、奉贤/嘉定（展柜）、阜城（集成房屋）。"
                  "做成一家就能顺着熟人网络复制。"),
        ("行业展会", "CHCC 医院建设大会、FBC 门窗幕墙、广交会、永康国际门业博览会、商业空间展 —— "
                 "展会抓的是正在花钱做市场、通常也在扩产的企业。"),
        ("招聘信息", "搜「MES 实施」「信息化主管」「数字化专员」。在招这类岗位 = 已有预算科目，"
                 "优先级直接从 B 提到 A，且对接人就是未来的项目负责人。"),
    ]:
        ws.write(r, 0, a, f["band"])
        ws.write(r, 1, b, f["note"])
        ws.set_row(r, 60)
        r += 1
    r += 1
    ws.write(r, 0, "免责与保密", f["sect"])
    r += 1
    ws.write(r, 0, "说明", f["band"])
    ws.write(r, 1, "全部信息来自公开渠道（企业官网、上市公司公告、政府/招标公告、工商公示、权威媒体、行业协会）；"
                   "只做初筛与排序，不是尽调结论；所有「失败面」都是待核实项，不构成对任何企业的负面判断；"
                   "联系方式为企业公开电话与公共邮箱，电话前请自行确认对方身份。"
                   "本文件所在仓库为公开仓库，用于商业开发请先移入私有仓库。", f["note"])
    ws.set_row(r, 70)
    return ws


def main():
    rows = list(csv.reader(open(CSV, encoding="utf-8-sig")))
    wb = xlsxwriter.Workbook(OUT)
    f = fmt(wb)
    sheet_list(wb, f, rows)
    sheet_verify(wb, f)
    sheet_rubric(wb, f)
    sheet_howto(wb, f)
    wb.close()
    print("写入", OUT, f"{os.path.getsize(OUT)/1024:.1f} KB")


if __name__ == "__main__":
    main()
