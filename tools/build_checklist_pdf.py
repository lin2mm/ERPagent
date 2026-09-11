"""生成 A4 打印版《目标客户筛选清单》。

与项目里其他产物一样：内容用中文，文件名用 ASCII。
A4 @200dpi = 1654 × 2339 px，可直接打印成 3 页表格，现场用笔填。
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PIL import Image, ImageDraw, ImageFont
from deckkit import T

W, H = 1654, 2339          # A4 @200dpi
M = 88                     # 页边距
CW = W - 2 * M             # 内容宽度
TOP, BOT = 96, 96

REG = "/tmp/fonts/NotoSansSC-Regular.otf"
BLD = "/tmp/fonts/NotoSansSC-Bold.otf"
_c = {}


def f(size, bold=False):
    k = (size, bold)
    if k not in _c:
        _c[k] = ImageFont.truetype(BLD if bold else REG, size)
    return _c[k]


def wrap(text, font, max_w):
    lines, cur = [], ""
    for ch in text:
        if ch == "\n":
            lines.append(cur); cur = ""
            continue
        if font.getlength(cur + ch) <= max_w:
            cur += ch
        else:
            lines.append(cur); cur = ch
    if cur:
        lines.append(cur)
    return lines


class Doc:
    def __init__(self):
        self.pages = []
        self._new()

    def _new(self):
        self.img = Image.new("RGB", (W, H), "#FFFFFF")
        self.d = ImageDraw.Draw(self.img)
        self.d.rectangle([0, 0, W, 14], fill="#" + T["accent"])
        self.y = TOP
        self.pages.append(self.img)

    def room(self, h):
        if self.y + h > H - BOT:
            self._new()

    def txt(self, x, y, text, size=24, bold=False, color=None, max_w=None):
        color = color or T["ink"]
        fnt = f(size, bold)
        lines = wrap(text, fnt, max_w or CW)
        lh = int(size * 1.45)
        for i, ln in enumerate(lines):
            self.d.text((x, y + i * lh), ln, font=fnt, fill="#" + color)
        return len(lines) * lh

    def h1(self, title, sub=None):
        h = 62
        self.txt(M, self.y, title, size=44, bold=True, color=T["ink"], max_w=CW)
        self.y += 62
        if sub:
            self.y += self.txt(M, self.y, sub, size=22, color=T["muted"], max_w=CW) + 4
        self.d.line([M, self.y + 6, W - M, self.y + 6], fill="#" + T["line"], width=2)
        self.y += 24

    def bar(self, text, color=None, note=None):
        color = color or T["accent"]
        self.room(76)
        self.d.rounded_rectangle([M, self.y, W - M, self.y + 52], radius=10, fill="#" + color)
        self.d.text((M + 20, self.y + 10), text, font=f(27, True), fill="#FFFFFF")
        if note:
            fnt = f(20)
            tw = fnt.getlength(note)
            self.d.text((W - M - 20 - tw, self.y + 15), note, font=fnt, fill="#E8EEF8")
        self.y += 68

    def check(self, text, tag=None, size=23, indent=0, color=None):
        """一行带方框的条目；tag 显示在右侧（如 __/5）。"""
        fnt = f(size)
        x = M + 8 + indent
        box = 26
        lines = wrap(text, fnt, CW - 60 - indent - (150 if tag else 0))
        lh = int(size * 1.42)
        self.room(lh * len(lines) + 14)
        self.d.rectangle([x, self.y + 3, x + box, self.y + 3 + box], outline="#" + T["muted"], width=2)
        for i, ln in enumerate(lines):
            self.d.text((x + box + 14, self.y + i * lh), ln, font=fnt,
                        fill="#" + (color or T["body"]))
        if tag:
            ft = f(21, True)
            tw = ft.getlength(tag)
            self.d.text((W - M - 8 - tw, self.y + 2), tag, font=ft, fill="#" + T["accent"])
        self.y += lh * len(lines) + 14

    def row(self, name, desc, tag=None, desc_size=21):
        """左标题 / 中说明 / 右得分 的三列行。"""
        fnt = f(desc_size)
        lw = 250
        lines = wrap(desc, fnt, CW - lw - 190)
        lh = int(desc_size * 1.42)
        h = max(lh * len(lines), 34) + 16
        self.room(h)
        self.d.text((M + 6, self.y + 1), name, font=f(23, True), fill="#" + T["ink"])
        for i, ln in enumerate(lines):
            self.d.text((M + lw, self.y + i * lh), ln, font=fnt, fill="#" + T["body"])
        if tag:
            ft = f(21, True)
            tw = ft.getlength(tag)
            self.d.text((W - M - 12 - tw, self.y + 1), tag, font=ft, fill="#" + T["accent"])
        self.y += h
        self.d.line([M, self.y - 6, W - M, self.y - 6], fill="#" + T["soft2"], width=1)

    def note(self, text, color=None, fill=None, size=22):
        color = color or T["body"]
        fill = fill or T["soft"]
        fnt = f(size)
        lines = wrap(text, fnt, CW - 40)
        lh = int(size * 1.42)
        h = lh * len(lines) + 28
        self.room(h)
        self.d.rounded_rectangle([M, self.y, W - M, self.y + h], radius=10, fill="#" + fill)
        for i, ln in enumerate(lines):
            self.d.text((M + 20, self.y + 14 + i * lh), ln, font=fnt, fill="#" + color)
        self.y += h + 16

    def sub(self, text, color=None, size=24):
        self.room(40)
        self.y += self.txt(M + 6, self.y, text, size=size, bold=True,
                           color=color or T["ink"], max_w=CW) + 8

    def blank(self, h=54):
        self.room(h)
        self.y += h

    def finish(self):
        n = len(self.pages)
        for i, img in enumerate(self.pages, 1):
            d = ImageDraw.Draw(img)
            ft = f(19)
            d.line([M, H - 74, W - M, H - 74], fill="#" + T["line"], width=1)
            d.text((M, H - 62), "目标客户筛选清单 · 打印填写 · 每接触一家填一份",
                   font=ft, fill="#" + T["muted"])
            s = f"{i} / {n}"
            d.text((W - M - ft.getlength(s), H - 62), s, font=ft, fill="#" + T["muted"])


doc = Doc()

# ───────────────────────── 标题 ─────────────────────────
doc.h1("目标客户筛选清单",
       "先过「一票否决」→ 企业侧与老板侧各打 35 分 → 定级 → 决定投入多少售前资源。满分 70 分。")

# ───────────────────────── 0 一票否决 ─────────────────────────
doc.bar("第 0 步 · 一票否决（任一不过 → 直接放弃或搁置）", T["red"])
doc.check("能见到拍板的人：股权清晰、创始人或家族实际控制；不是只跟 IT、生产部长谈。")
doc.check("近 12 个月没有失败的信息化项目；若有，老板能说清原因，且愿意听我们怎么回应。")
doc.check("业务负责人（生产 / 设计 / 财务）若反对，老板明确表示会处理。")
doc.check("有数据载体：在用进销存 / ERP / 总账之一，且有物料编码（编码乱没关系）。")
doc.check("有现金：预算几十万至几百万，且不影响正常经营 —— 不靠贷款、不靠赊账做项目。")
doc.check("一把手愿意亲自下场：能参加双周演示、能安排关键用户工时。（这是最强的筛选器）")
doc.blank(10)

# ───────────────────────── 1 企业侧 ─────────────────────────
doc.bar("第 1 步 · 企业侧打分（每项 0–5 分，满分 35）", T["green"])
doc.row("E1 规模匹配", "0 分：营收 < 5000 万，员工 < 80 人　|　3 分：5000 万–1 亿　|　"
                      "5 分：1–10 亿、80–800 人（更大规模慎做：你从方案方变成分包商）", "____ / 5")
doc.row("E2 痛点具体性", "0 分：只说\"该数字化了\"　|　3 分：能说出现象　|　"
                        "5 分：能说出具体事件与金额（被罚过 / 压了几百万 / 赔过 / 丢过标）", "____ / 5")
doc.row("E3 外部压力", "0 分：没有　|　3 分：正在计划　|　"
                      "5 分：已有硬要求（大客户溯源、招标资质、合规检查、资本化）", "____ / 5")
doc.row("E4 数据基础", "0 分：全手工且无人愿清洗　|　3 分：有进销存但编码混乱　|　"
                      "5 分：有编码，且有打标 / 条码设备", "____ / 5")
doc.row("E5 业务稳定性", "0 分：在生死线上或剧烈转型　|　3 分：行业平淡但稳定　|　"
                        "5 分：订单饱满、正在扩产", "____ / 5")
doc.row("E6 付费能力", "0 分：要靠贷款或赊账　|　3 分：能付但会反复压价　|　"
                      "5 分：有预算科目、接受按里程碑付款", "____ / 5")
doc.row("E7 行业价值密度", "0 分：大批量标准品，拼设备拼自动化　|　3 分：多品种小批量　|　"
                          "5 分：非标定制 / 原材料占比高 / 有追溯要求", "____ / 5")
doc.note("企业侧小计：____ / 35", T["ink"], T["green_l"], 24)
doc.blank(6)

# ───────────────────────── 2 老板侧 ─────────────────────────
doc.bar("第 2 步 · 老板侧打分（每项 0–5 分，满分 35）", T["accent"])
doc.row("B1 亲自下场意愿 ★", "0 分：见不到人、全权委托　|　3 分：愿意出席关键会议　|　"
                            "5 分：愿意把关键用户按在会议室签字、参加双周演示（最强筛选器）", "____ / 5")
doc.row("B2 说得出瓶颈数字", "0 分：说不清不良率 / 周转 / 毛利　|　3 分：说得大概　|　"
                            "5 分：报得出具体数字，且知道卡在哪道工序", "____ / 5")
doc.row("B3 数字化体感", "0 分：排斥电子设备　|　3 分：会用微信、钉钉　|　"
                        "5 分：手机上看数据，主动问\"能不能手机上批\"", "____ / 5")
doc.row("B4 吃过亏且记得住", "0 分：没吃过亏或记不住　|　3 分：知道有问题但没算过账　|　"
                            "5 分：能讲出一次具体的损失事故", "____ / 5")
doc.row("B5 扩张野心", "0 分：守成、准备卖厂　|　3 分：稳住就好　|　"
                      "5 分：要接大客户 / 扩产 / 做品牌（才愿为管理能力付费）", "____ / 5")
doc.row("B6 年龄与精力", "0 分：55 岁以上且无接棒人　|　3 分：55 岁以上但有二代 / 职业经理人　|　"
                        "5 分：35–55 岁、仍在一线管事", "____ / 5")
doc.row("B7 决策效率", "0 分：三次沟通无进展、反复\"再看看\"　|　3 分：需内部商量但有节奏　|　"
                      "5 分：当场能定方向、指定对接人", "____ / 5")
doc.note("老板侧小计：____ / 35", T["ink"], T["accent_l"], 24)
doc.blank(6)

# ───────────────────────── 3 负面清单 ─────────────────────────
doc.bar("第 3 步 · 负面清单（命中任一条扣 5 分；带 ★ 的直接放弃）", T["orange"])
for i, t in enumerate([
    "老板不在现场，全权交给职业经理人 ★",
    "只问\"多少钱\"，不问\"怎么做\" ★（你只是他的比价工具）",
    "要\"先做个小样看看\"，却不肯投入任何资源",
    "老板认为自己什么都懂，要你\"按我说的做\"",
    "业务负责人明确反对，而老板不打算处理 ★",
    "数据极度混乱，且没人愿意清洗 ★",
    "家族内斗 / 股权纠纷，决策随时被推翻 ★",
    "行业正在衰退，或企业准备卖掉",
    "近 12 个月刚有失败的信息化项目，且说不清原因",
    "要你先做完整方案再谈钱（大概率拿去给别家报价）★",
], 1):
    doc.check(t, size=22)
doc.blank(6)

# ───────────────────────── 4 定级 ─────────────────────────
doc.bar("第 4 步 · 定级与打法（总分 = 企业侧 + 老板侧，满分 70）", T["purple"])
doc.row("S 级 · ≥ 56 分", "立即投入：安排同行业活案例参观 → 4 周内做出\"只读驾驶舱\" → 直接谈分期与里程碑。",
        "全力")
doc.row("A 级 · 42–55 分", "重点培育：先做半天走车间 + 1 页症断书，建立信任后再谈主线。", "中等")
doc.row("B 级 · 28–41 分", "保持观察：定期发行业案例、邀请参加活动，等外部压力出现。", "低")
doc.row("C 级 · < 28 分", "礼貌退出；命中一票否决或 ★ 的，同样按 C 级处理，不投入售前资源。", "零")
doc.note("两条修正规则：① 总分高但老板不愿亲自下场 → 按 B 级处理；"
         "② 总分中等但老板肯下场、且有外部压力 → 可按 A 级投入。", T["ink"], T["amber_l"])

# ───────────────────────── 5 必问 12 问 ─────────────────────────
doc.bar("第 5 步 · 见面必问的 12 个问题（用来填上面的分）", T["ink"])
doc.sub("关于钱", T["green"])
doc.check("去年哪三笔钱最让您心疼？大概各是多少？", size=22)
doc.check("单张订单赚不赚钱，您什么时候能知道？", size=22)
doc.check("这次投入大概什么量级，走什么审批流程？", size=22)
doc.sub("关于人", T["accent"])
doc.check("这事谁能拍板？除了您，还有谁的意见很关键？", size=22)
doc.check("生产、设计、财务的负责人，您觉得他们会怎么想？", size=22)
doc.check("上次搞系统是谁在推？最后为什么停了？", size=22)
doc.sub("关于事", T["orange"])
doc.check("从接单到发货，最常卡在哪一步？一般卡多久？", size=22)
doc.check("客户投诉最多的是哪一类？", size=22)
doc.check("现在有没有客户要求溯源，或要求提供台账？", size=22)
doc.sub("关于风险", T["red"])
doc.check("如果三个月看不到东西，您能接受吗？", size=22)
doc.check("上线期间，您最怕影响到什么？", size=22)
doc.check("数据这块，您自己觉得乱到什么程度？", size=22)

# ───────────────────────── 6 走车间 30 分钟 ─────────────────────────
doc.bar("第 6 步 · 走车间 30 分钟：看这 8 个信号", T["accent_d"])
doc.check("排产用的是白板或纸质单 → 管理全靠人，系统有空间")
doc.check("货架上有积灰的原料 / 卷料 → 呆滞资金，账就在眼前")
doc.check("工人说不出手上这批货给哪个客户 → 追溯意识弱，也是机会")
doc.check("班组长用本子记数 → 计件争议的源头")
doc.check("有闲置的大屏或扫码枪 → 上次失败的遗迹，同时是现成资产")
doc.check("打标 / 喷码设备在跑，但数据没进系统 → 最值钱的切入点")
doc.check("问\"上个月报废多少\" → 答不上来，说明账不清楚")
doc.check("墙上还有某个系统的旧海报 → 上次用了多久？现在谁在用？")

# ───────────────────────── 7 结论记录 ─────────────────────────
doc.bar("第 7 步 · 结论记录", T["green"])
doc.blank(6)
for line in ["企业名称：______________________________　"
             "接触人 / 职务：______________________________　"
             "日期：____ 年 __ 月 __ 日",
             "一票否决：□ 通过　　□ 不通过（第 ______ 条）　　命中 ★：______ 条",
             "企业侧 ______ / 35　　老板侧 ______ / 35　　总分 ______ / 70",
             "定级：□ S　□ A　□ B　□ C　　下一步动作：____________________________________",
             "最大风险（一句话）：____________________________________________________________",
             "下次接触时间：________________　负责人：________________"]:
    doc.room(48)
    doc.y += doc.txt(M + 6, doc.y, line, size=23, color=T["body"], max_w=CW) + 18

doc.finish()

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "pdf",
                   "CHECKLIST-Customer-Screening.pdf")
os.makedirs(os.path.dirname(out), exist_ok=True)
imgs = doc.pages
imgs[0].save(out, "PDF", resolution=200, save_all=True, append_images=imgs[1:])
for i, im in enumerate(imgs, 1):
    im.save(f"/tmp/checklist-p{i}.png")
print(f"页数 {len(imgs)}  →  {out}  ({os.path.getsize(out)/1024:.0f} KB)")
print("预览:", ", ".join(f"/tmp/checklist-p{i}.png" for i in range(1, len(imgs) + 1)))
