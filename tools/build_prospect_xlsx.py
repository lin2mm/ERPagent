#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""由 prospects/Target-Companies.csv 生成可下载的 Excel 名单（v3 · 含重点区域）。

    /tmp/v/bin/python tools/build_prospect_xlsx.py

产出 prospects/Target-Companies.xlsx，五张表：
    重点区域速览      —— 长三角 / 珠三角 单独一张，按分数排序，看这一张就够
    名单              —— 44 家 × 24 列（含官网核实、联系方式、社媒、触达路径）
    本轮核实纪要      —— 核实方法、结果统计、修正清单 + 本轮新增区域客户
    评分口径          —— E1–E7 / B1–B7 / 定级线 / 修正规则
    使用说明与获客渠道

CSV 是唯一数据源：改名单改 CSV（或用 tools/build_targets.py 重新组装），再跑本脚本。
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
REGION_STYLE = {"长三角": "#0E6B6B", "珠三角": "#8A5A0B"}

# 24 列宽度
WIDTHS = [6, 32, 12, 26, 8, 11, 24, 38, 42, 22, 18, 26, 26, 42, 12, 12, 8, 7, 34, 34, 26, 26, 9, 10]
WRAP_COLS = {1, 3, 7, 8, 9, 10, 11, 12, 13, 14, 15, 18, 19, 20, 21}
CENTER_COLS = {0, 2, 4, 5, 16, 17, 22, 23}


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
        good=wb.add_format({"font_size": 10, "font_color": "#1E6F50", "text_wrap": True,
                            "valign": "top"}),
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
        rfmt = wb.add_format({"font_size": 9.5, "bold": True,
                              "font_color": REGION_STYLE.get(row[2], MUTED),
                              "align": "center", "valign": "vcenter",
                              "border": 1, "border_color": LINE})
        for c, val in enumerate(row):
            cell = f["cell"]
            if c == 0:
                cell = tfmt
            elif c == 2:
                cell = rfmt
            elif c == 1:
                cell = f["adv"]
            elif c == 5 and val.startswith("http"):
                ws.write_url(r, c, val, f["link"], val)          # 官网可点
                continue
            elif c in CENTER_COLS:
                cell = f["num"]
            if c == 10:                                            # 邮箱可点
                one = re.fullmatch(r"[\w.+-]+@[\w-]+\.[\w.]+", val.strip())
                if one:
                    ws.write_url(r, c, "mailto:" + val.strip(), f["link"], val.strip())
                    continue
            ws.write(r, c, val, cell)
        ws.set_row(r, 104)
    ws.freeze_panes(1, 3)
    ws.autofilter(0, 0, len(rows) - 1, len(head) - 1)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(0)
    return ws


def sheet_focus(wb, f, rows):
    """第一张表：只看长三角 / 珠三角。"""
    ws = wb.add_worksheet("重点区域速览")
    head24 = rows[0]
    idx = {h: i for i, h in enumerate(head24)}
    cols = [("企业名称", 30), ("区域", 9), ("城市", 12), ("行业细分", 14),
            ("适配案例", 9), ("初筛分(企业侧/35)", 8), ("优先级建议", 46),
            ("联系电话", 26), ("社媒（公众号 / 抖音 / 1688 / 其他）", 26),
            ("触达路径（细化到联系方式）", 46)]
    ws.set_column(0, 0, 5)
    for i, (_, w) in enumerate(cols, start=1):
        ws.set_column(i, i, w)
    ws.write(0, 0, "长三角 · 珠三角 重点客户（按初筛分排序）", f["title"])
    ws.write(1, 0, "这两片区域单列一页：长三角 19 家（上海 / 江苏 / 浙江）+ 珠三角 14 家（广州 / 佛山 / 东莞 / 深圳 / 江门 / 清远）。"
                   "分数是「企业侧」公开信息估算分（满分 35），老板侧 35 分必须面谈才能打。全量 44 家见「名单」表。", f["note"])
    ws.set_row(1, 30)
    for c, (name, _) in enumerate(cols, start=1):
        ws.write(3, 0, "#", f["band"])
        ws.write(3, c, name, f["band"])
    ws.set_row(3, 26)
    target = [r for r in rows[1:] if r[idx["区域"]] in ("长三角", "珠三角")]
    target.sort(key=lambda r: -float(r[idx["初筛分(企业侧/35)"]] or 0))
    for n, row in enumerate(target, start=1):
        r = 3 + n
        ws.write(r, 0, n, f["num"])
        for c, (name, _) in enumerate(cols, start=1):
            val = row[idx[name]]
            if name == "区域":
                cell = wb.add_format({"font_size": 9.5, "bold": True,
                                      "font_color": REGION_STYLE.get(val, MUTED),
                                      "align": "center", "valign": "top",
                                      "border": 1, "border_color": LINE})
            elif name in ("适配案例", "初筛分(企业侧/35)"):
                cell = f["num"]
            elif name == "优先级建议":
                cell = f["adv"]
            else:
                cell = f["cell"]
            ws.write(r, c, val, cell)
        ws.set_row(r, 62)
    ws.write(3 + len(target) + 2, 0,
             "怎么用：先按分数从上往下打 3 个电话（每次只问三句 —— 订单排到什么时候？有没有因交期/批次被罚过？老板自己管不管生产？），"
             "三句里两句是具体事件就能约「半天走车间」。", f["good"])
    ws.write(3 + len(target) + 3, 0,
             "注意：分数只反映公开信息，不代表企业优劣；「待核实」项是去现场要问的问题，不是结论。", f["note"])
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
    ("山东省博兴县聚鑫源", "官网另给「博兴经济开发区富源二路 96 号」+ 外贸部专线，与名单「店子工业园聚鑫源路 6 号」不一致 → 拜访前确认厂区"),
]

NEW_THIS_ROUND = [
    ("长三角", "S", "27", "江苏赛康医疗设备（张家港）", "新三板 870098、290 人、5 万㎡：出口 100+ 国的批次溯源刚需"),
    ("长三角", "A", "26", "江苏大力神科技（丹阳）", "900 亩、8 条镀锌线 + 彩涂线：与案例 A 同工艺"),
    ("长三角", "A", "26", "苏州德品医疗（苏州高新区）", "500+ 三甲医院、智慧护理系统：自带信息化语言"),
    ("长三角", "A", "25", "江苏丽岛新材料（常州）", "彩涂铝卷、年销 13 亿、员工 720：辊涂工艺与案例 A 同构"),
    ("长三角", "A", "24", "名耀家具（苏州相城）", "注册 8000 万、4 万㎡、300+ 人：酒店整店交付"),
    ("长三角", "B", "23", "浙江神将门业（武义）", "注册 8000 万、8 条线 100 万樘：渠道 + 工程双台账"),
    ("长三角", "B", "22", "南京布尔特医疗（南京）", "参编医院建设指南；生产在分支机构，规模待核"),
    ("长三角", "C", "19", "上海玄策展览（上海）", "4500㎡ 工厂、85 人：展会短周期，作轻量样板"),
    ("珠三角", "S", "27", "广州博生医疗家具（花都）", "注册 1.08 亿、600+ 人、10 万㎡：投标驱动型工厂"),
    ("珠三角", "A", "26", "广东华展家具 · 华展医疗（清远）", "注册 5200 万、500+ 医疗机构：参编医院建设指南"),
    ("珠三角", "A", "25", "东莞领展展示用品（桥头）", "2.5 万㎡、301–500 人、出口 50+ 国：Sedex 验厂要追溯"),
    ("珠三角", "A", "24", "广东晟麟展览展示 · 铁衣卫（沙田）", "6 万㎡、500+ 技工：手机品牌展柜 + 精密钣金"),
    ("珠三角", "B", "23", "广州市南粤防火门（从化/白云）", "1991 年老厂、4 基地、30 万樘：3C 与验收台账"),
    ("珠三角", "B", "22", "广东德云医养家具（佛山南海）", "医养双线、有展厅；厂房与人数口径矛盾待核"),
    ("珠三角", "B", "21", "广东皓晖消防设备（佛山南海）", "不锈钢防火门、注册 600 万：规模小、决策快"),
    ("珠三角", "C", "20", "深圳昊晟展示工程（坪山）", "珠宝/化妆品展柜；公开信息薄，作信息源"),
]


def sheet_verify(wb, f):
    ws = wb.add_worksheet("本轮核实纪要")
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 100)
    ws.write(0, 0, "官网反向验证 · 两轮纪要", f["title"])
    ws.write(1, 0, "方法：只看企业官网（首页/关于我们/产品/联系方式）、工商公示与公开披露；"
                   "官网写的口径优先于黄页与推广软文；核不上的一律写「待核实」，不做断言。", f["note"])
    ws.set_row(1, 30)
    ws.write(3, 0, "名单现状", f["sect"])
    stats = [("覆盖企业", "44 家（长三角 19 · 珠三角 14 · 其他 11）"),
             ("拿到官网链接", "33 家"), ("拿到企业电话", "41 家"),
             ("拿到社媒入口", "26 家（公众号 / 抖音 / 微博 / 1688 / 视频号）"),
             ("本轮新增", "16 家（长三角 8 + 珠三角 8，全部完成官网反向验证与联系方式细化）"),
             ("结论变化（上一轮）", "3 家升档（益德 / 香乡 / 大兆）、5 家降档、2 家建议移出")]
    r = 4
    for a, b in stats:
        ws.write(r, 0, a, f["band"])
        ws.write(r, 1, b, f["note"])
        ws.set_row(r, 20)
        r += 1
    r += 1
    ws.write(r, 0, "本轮新增（重点区域）", f["sect"])
    r += 1
    for h, w in (("区域", 10), ("优先级", 8), ("初筛分", 8), ("企业", 34), ("一句话", 70)):
        ws.write(r, 0 if h == "区域" else (1 if h == "优先级" else (2 if h == "初筛分" else (3 if h == "企业" else 4))), h, f["band"])
    ws.set_column(2, 2, 8)
    ws.set_column(3, 3, 34)
    ws.set_column(4, 4, 70)
    r += 1
    for region, tier, score, name, note in NEW_THIS_ROUND:
        ws.write(r, 0, region, f["note"])
        ws.write(r, 1, tier, f["note"])
        ws.write(r, 2, score, f["note"])
        ws.write(r, 3, name, f["note"])
        ws.write(r, 4, note, f["note"])
        ws.set_row(r, 18)
        r += 1
    r += 1
    ws.write(r, 0, "上一轮修正清单", f["sect"])
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
    r += 2
    ws.write(r, 0, "社媒怎么用", f["band"])
    ws.write(r, 1, "公众号 / 视频号 / 抖音不是拿来群发广告的：先看它最近 3 个月发什么 —— "
                   "发订单和产能说明「缺单」，发资质与中标说明「在投标」，发设备与招聘说明「在扩产」。"
                   "看出来再打电话，第一句话就能说到点上。没采到社媒的写「待补」，下一轮补。", f["note"])
    ws.set_row(r, 60)


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
            ("S 级 ≥ 56（总分）", "", "立即投入：同行业活案例参观 → 4 周内做出只读驾驶舱 → 谈分期与里程碑"),
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
            ("官网核实的作用", "", "先去掉名字与网址对不上的、规模超出的、主体查不到的，再决定投入顺序"),
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
                                        "对不上就先降到 B 级观察。官网打不开的看社媒最近三个月在发什么。"),
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
    ws.write(r, 0, "区域打法（本轮重点）", f["sect"])
    r += 1
    for a, b in [
        ("长三角", "上海（展柜/系统家具）+ 苏州—张家港（医疗设备+家具）+ 常州/丹阳（彩涂与涂镀）："
                 "一天能连访 2–3 家，且客户之间互相认识 —— 做成一家就能顺着同城圈子复制。"),
        ("珠三角", "广州花都（医疗家具三家可连访）+ 佛山南海（医养/门窗）+ 东莞（展柜与钣金）+ 深圳（展柜）："
                 "同一个上午就能从花都跑到佛山；出口型企业多，验厂与批次追溯是天然切入点。"),
        ("怎么用区域列", "「名单」表按区域分了三档：长三角 / 珠三角 / 其他 · 省市。先做前两档，"
                     "其他档（山东彩涂、河北门业、天津彩板等）作为出差顺路或跨区复制用。"),
    ]:
        ws.write(r, 0, a, f["band"])
        ws.write(r, 1, b, f["note"])
        ws.set_row(r, 58)
        r += 1
    r += 1
    ws.write(r, 0, "四种持续获客渠道", f["sect"])
    r += 1
    for a, b in [
        ("招标中标名单", "ccgp.gov.cn 与各省招标平台，搜「家具」「防火门」「彩涂板」。中标公告里有甲方、乙方、金额、工期 —— "
                    "工期违约条款抄下来，就是方案第一页的「账」。"),
        ("产业带集群", "长三角：上海奉贤（展柜）、苏州—张家港（医疗）、常州丹阳（涂镀）、武义永康（门业）；"
                  "珠三角：广州花都（医疗家具）、佛山南海（医养门窗）、东莞（展示道具与钣金）、深圳（珠宝展柜）。"),
        ("行业展会", "CHCC 医院建设大会、FBC 门窗幕墙、广交会、永康国际门业博览会、深圳国际家具展、商业空间展 —— "
                 "展会抓的是正在花钱做市场、通常也在扩产的企业；医疗家具与门业的展会还能一次见全一个行业。"),
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
                   "只做初筛与排序，不是尽调结论；所有「失败面」与「待核实」都是现场要问的问题，不构成对任何企业的负面判断；"
                   "联系方式为企业公开电话与公共邮箱（部分为业务员手机，来自官网或公开商铺），联系前请自行确认对方身份，"
                   "并遵守个人信息保护相关要求；本文件所在仓库为公开仓库，用于商业开发请先移入私有仓库。", f["note"])
    ws.set_row(r, 78)
    return ws


def main():
    rows = list(csv.reader(open(CSV, encoding="utf-8-sig")))
    assert len(rows[0]) == 24, f"CSV 应为 24 列，当前 {len(rows[0])} 列（先跑 tools/build_targets.py）"
    wb = xlsxwriter.Workbook(OUT)
    f = fmt(wb)
    sheet_focus(wb, f, rows)
    sheet_list(wb, f, rows)
    sheet_verify(wb, f)
    sheet_rubric(wb, f)
    sheet_howto(wb, f)
    wb.close()
    print(f"写入 {OUT}（{len(rows) - 1} 家 × {len(rows[0])} 列 · {os.path.getsize(OUT)/1024:.1f} KB）")


if __name__ == "__main__":
    main()
