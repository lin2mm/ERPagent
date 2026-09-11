"""售前说服版（老板决策简报）共用版式。

同一套骨架服务两个案例：
    先讲客户的账 → 讲买到什么 → 讲担心什么 → 讲不做什么 → 讲要客户出什么
    → 讲多少钱 → 讲多久见效 → 讲怎么验收 → 讲适不适合 → 讲下一步。

设计依据见 docs/07-Boss-Perspective-Sales-Strategy.md：
  · 老板买的不是系统，是「少赔钱 / 多接单 / 不用天天盯」
  · 最大的顾虑不是价格，是「我推不动」—— 所以有一页专讲推行机制
  · 主动减承诺（不做什么）是最强的信任建立动作
  · 把第一次投入和第一次风险做小，是提高成功率最有效的一招
  · 所有金额只给「算法与口径」，不给假精确的承诺值

案例差异只体现在内容字典里，版式完全共用 —— 这样每来一个新客户，
只要填一份 content 就能出一份老板版简报。
"""

from deckkit import Deck, T, L, SW, PAD

X3 = 0.62
FULL = 12.09                      # 内容总宽（0.62 → 12.71）
CW4 = (FULL - 0.84) / 4           # 四列卡片宽
GAP4 = 0.28


# ───────────────────────────── 通用小部件 ─────────────────────────────
def _pain_cards(d, s, cards, y=2.06, h=2.34):
    """四张痛点卡：标签 + 标题 + 说明。"""
    for i, (tag, title, body, col) in enumerate(cards):
        x = X3 + i * (CW4 + GAP4)
        d.rect(s, x, y, CW4, h, fill=T["white"], line=T["line"], radius=0.10)
        d.rect(s, x, y, CW4, 0.055, fill=col, radius=0.0)
        d.text(s, x + 0.20, y + 0.17, CW4 - 0.40, 0.26,
               [L(tag, size=9.6, bold=True, color=col)])
        d.text(s, x + 0.20, y + 0.46, CW4 - 0.40, 0.66,
               [L(title, size=14.5, bold=True, color=T["ink"])])
        d.text(s, x + 0.20, y + 1.18, CW4 - 0.40, h - 1.32,
               [L(body, size=10.2, color=T["body"])])


def _rowlist(d, s, rows, y=2.06, h=0.78, gap=0.09, lw=2.70, x=X3, w=FULL,
             fill=None, title_col=None, body_col=None, bar_col=None):
    """整幅行列表：左侧粗体标题 + 右侧说明。"""
    cy = y
    for title, body, col in rows:
        d.rect(s, x, cy, w, h, fill=fill or T["soft"], line=T["line"], radius=0.08)
        d.rect(s, x, cy, 0.075, h, fill=bar_col or col, radius=0.0)
        d.text(s, x + 0.22, cy + 0.11, lw, h - 0.22,
               [L(title, size=11.8, bold=True, color=title_col or T["ink"])])
        d.text(s, x + 0.22 + lw + 0.06, cy + 0.11, w - (0.22 + lw + 0.06) - 0.20,
               h - 0.22, [L(body, size=10.4, color=body_col or T["body"])])
        cy += h + gap
    return cy


def _darkbar(d, s, y, h, text, size=11.4, x=X3, w=FULL):
    d.rect(s, x, y, w, h, fill=T["ink"], radius=0.10)
    d.text(s, x + 0.30, y + 0.10, w - 0.60, h - 0.20,
           [L(text, size=size, color=T["white"])])


def _pair_cards(d, s, pairs, y=2.06, h=2.30, gap=0.30, say_size=11.6, ans_size=10.3):
    """2×2 的「老板的话 → 我们怎么答」卡片。"""
    w = (FULL - gap) / 2
    for i, (say, ans, col) in enumerate(pairs):
        x = X3 + (i % 2) * (w + gap)
        yy = y + (i // 2) * (h + 0.10)
        d.rect(s, x, yy, w, h, fill=T["white"], line=T["line"], radius=0.10)
        d.rect(s, x, yy, 0.075, h, fill=col, radius=0.0)
        d.text(s, x + 0.26, yy + 0.16, w - 0.52, 0.32,
               [L(say, size=say_size, bold=True, color=col)])
        d.text(s, x + 0.26, yy + 0.58, w - 0.52, h - 0.72,
               [L(ans, size=ans_size, color=T["body"])])


# ───────────────────────────── 十六页骨架 ─────────────────────────────
def slide_cover(d, c):
    s = d.slide(dark=True, notes=c["notes"])
    d.text(s, PAD, 1.20, 9.6, 0.32,
           [L(c["kicker"], size=12.0, bold=True, color=T["accent"])])
    d.text(s, PAD, 1.58, 11.4, 1.55,
           [L(c["title"], size=40, bold=True, color=T["white"]),
            L(c["sub"], size=16.5, color="A9BAD4", space=8)])
    d.hline(s, PAD, 3.30, 5.6, T["accent"], 2.0)
    d.text(s, PAD, 3.54, 11.4, 1.20, [L(c["lead"], size=12.6, color="C6D2E4")])
    for i, (v, lab) in enumerate(c["meta"]):
        x = PAD + i * 2.30
        d.rect(s, x, 4.94, 2.06, 0.86, fill="1B2C46", line="2C405F", radius=0.10)
        d.text(s, x, 5.04, 2.06, 0.42,
               [L(v, size=19, bold=True, color=T["white"], align="c")])
        d.text(s, x, 5.46, 2.06, 0.26,
               [L(lab, size=9.5, color="8FA3BF", align="c")])
    d.card(s, PAD, 6.06, FULL, 0.76, fill="1B2C46", lines=[
        L(c["decl"], size=10.0, color="A9BAD4")])


def slide_bills(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    _pain_cards(d, s, c["cards"], h=2.34)
    d.card(s, X3, 4.56, FULL, 1.16, title=c["trans_title"], fill=T["accent_l"],
           accent=T["accent"], title_size=12.2, lines=[
               L(c["trans_body"], size=10.6, color=T["ink"])])
    _darkbar(d, s, 5.86, 0.92, c["bar"])


def slide_three(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    w = (FULL - 2 * 0.30) / 3
    for i, (title, body, col) in enumerate(c["cards"]):
        x = X3 + i * (w + 0.30)
        d.rect(s, x, 2.06, w, 2.40, fill=T["white"], line=T["line"], radius=0.10)
        d.rect(s, x, 2.06, w, 0.055, fill=col, radius=0.0)
        d.text(s, x + 0.24, 2.28, w - 0.48, 0.40,
               [L(title, size=17, bold=True, color=T["ink"])])
        d.text(s, x + 0.24, 2.86, w - 0.48, 1.42,
               [L(body, size=11.0, color=T["body"])])
    d.card(s, X3, 4.62, FULL, 2.06, title=c["foot_title"], fill=T["soft"],
           accent=T["ink"], title_size=12.5, lines=[
               L(x, size=11.0, color=T["ink"], space=4) for x in c["foot"]])


def slide_gifts(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    for i, (icon, title, body) in enumerate(c["items"]):
        x = X3 + (i % 4) * (CW4 + GAP4)
        y = 2.02 + (i // 4) * 2.36
        d.rect(s, x, y, CW4, 2.24, fill=T["white"], line=T["line"], radius=0.10)
        d.rect(s, x, y, CW4, 0.05, fill=T["accent"], radius=0.0)
        d.text(s, x + 0.18, y + 0.16, CW4 - 0.36, 0.32,
               [L(f"{icon}  {title}", size=12.6, bold=True, color=T["ink"])])
        d.text(s, x + 0.18, y + 0.60, CW4 - 0.36, 1.50,
               [L(body, size=10.0, color=T["body"])])


def slide_money(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    d.table(s, X3, 2.06, FULL, c["rows"], [0.22, 0.36, 0.42],
            size=10.2, row_h=0.62)
    d.card(s, X3, 5.98, FULL, 0.82, fill=T["green_l"], accent=T["green"],
           lines=[L(c["foot"], size=10.6, color=T["ink"])], title=None)


def slide_fears(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    _pair_cards(d, s, c["pairs"])


def slide_people(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    d.table(s, X3, 2.06, FULL, c["rows"], [0.20, 0.36, 0.44],
            size=10.2, row_h=0.84)
    _darkbar(d, s, 6.00, 0.82, c["bar"])


def slide_notdo(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    _rowlist(d, s, c["rows"], y=2.06, h=0.74, gap=0.08, lw=2.55)
    d.text(s, X3, 6.32, FULL, 0.46,
           [L(c["foot"], size=10.6, bold=True, color=T["red"])])


def slide_need(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    w = (FULL - 2 * 0.30) / 3
    for i, (title, body, col) in enumerate(c["cards"]):
        x = X3 + i * (w + 0.30)
        d.rect(s, x, 2.06, w, 2.30, fill=T["white"], line=T["line"], radius=0.10)
        d.rect(s, x, 2.06, w, 0.055, fill=col, radius=0.0)
        d.text(s, x + 0.24, 2.26, w - 0.48, 0.40,
               [L(title, size=15.5, bold=True, color=T["ink"])])
        d.text(s, x + 0.24, 2.82, w - 0.48, 1.36,
               [L(body, size=10.6, color=T["body"])])
    d.card(s, X3, 4.52, FULL, 2.24, title=c["foot_title"], fill=T["soft"],
           accent=T["accent"], title_size=12.5, lines=[
               L(x, size=11.0, color=T["ink"], space=4) for x in c["foot"]])


def slide_budget(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    d.table(s, X3, 2.06, FULL, c["rows"], [0.16, 0.42, 0.42],
            size=10.2, row_h=0.60)
    w = (FULL - 2 * 0.26) / 3
    for i, (title, body, col) in enumerate(c["rules"]):
        x = X3 + i * (w + 0.26)
        d.card(s, x, 5.78, w, 1.06, title=title, accent=col, fill=T["white"],
               title_size=12.0, pad=0.14,
               lines=[L(body, size=9.8, color=T["body"])])


def slide_steps(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    y, h, gap = 2.06, 0.74, 0.07
    lw, rw = 3.05, 2.95
    for stage, when, what, stop in c["rows"]:
        d.rect(s, X3, y, FULL, h, fill=T["soft"], line=T["line"], radius=0.08)
        d.rect(s, X3, y, 0.075, h, fill=T["accent"], radius=0.0)
        d.text(s, X3 + 0.22, y + 0.10, lw - 0.20, 0.32,
               [L(stage, size=11.6, bold=True, color=T["ink"])])
        d.text(s, X3 + 0.22, y + 0.42, lw - 0.20, 0.24,
               [L(when, size=9.4, color=T["muted"])])
        d.text(s, X3 + 0.22 + lw, y + 0.11, FULL - lw - rw - 0.70, h - 0.22,
               [L(what, size=10.2, color=T["body"])])
        d.text(s, X3 + FULL - rw - 0.24, y + 0.11, rw, h - 0.22,
               [L(stop, size=10.0, bold=True, color=T["orange"])])
        y += h + gap
    _darkbar(d, s, y + 0.02, 0.68, c["foot"])


def slide_accept(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    w = (FULL - 2 * 0.30) / 3
    for i, (title, body, col) in enumerate(c["tests"]):
        x = X3 + i * (w + 0.30)
        d.rect(s, x, 2.06, w, 2.36, fill=T["white"], line=T["line"], radius=0.10)
        d.rect(s, x, 2.06, w, 0.055, fill=col, radius=0.0)
        d.text(s, x + 0.24, 2.26, w - 0.48, 0.40,
               [L(title, size=14.5, bold=True, color=T["ink"])])
        d.text(s, x + 0.24, 2.82, w - 0.48, 1.42,
               [L(body, size=10.6, color=T["body"])])
    d.card(s, X3, 4.58, FULL, 2.14, title=c["foot_title"], fill=T["red_l"],
           accent=T["red"], title_size=12.5, lines=[
               L(x, size=11.0, color=T["ink"], space=4) for x in c["foot"]])


def slide_fit(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    w = (FULL - 0.30) / 2
    for i, (head, items, col, fill) in enumerate(
            [(c["yes_head"], c["yes"], T["green"], T["green_l"]),
             (c["no_head"], c["no"], T["red"], T["red_l"])]):
        x = X3 + i * (w + 0.30)
        d.card(s, x, 2.06, w, 3.56, title=head, fill=fill, accent=col,
               title_size=13.5, lines=[L(t, size=10.8, color=T["ink"], space=6)
                                       for t in items])
    _darkbar(d, s, 5.78, 0.94, c["bar"])


def slide_push(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    _rowlist(d, s, c["rows"], y=2.06, h=0.95, gap=0.09, lw=3.30)
    _darkbar(d, s, 6.20, 0.64, c["bar"])


def slide_next(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], sub=c["sub"], notes=c["notes"])
    _rowlist(d, s, c["rows"], y=2.02, h=0.90, gap=0.09, lw=2.20)
    w = (FULL - 0.30) / 2
    for i, (title, lines) in enumerate(c["prep"]):
        x = X3 + i * (w + 0.30)
        d.card(s, x, 5.12, w, 1.68, title=title, accent=T["accent"] if i == 0 else T["green"],
               fill=T["white"], title_size=12.2, pad=0.16,
               lines=[L(t, size=10.0, color=T["body"], space=3) for t in lines])


def slide_cheat(d, c):
    s = d.slide(kicker=c["kicker"], title=c["title"], notes=c["notes"], dark=True)
    y, h, gap = 1.88, 0.68, 0.07
    for q, a in c["rows"]:
        d.rect(s, X3, y, FULL, h, fill="1B2C46", line="2C405F", radius=0.08)
        d.text(s, X3 + 0.24, y + 0.10, 2.35, h - 0.20,
               [L(q, size=11.4, bold=True, color=T["accent"])])
        d.text(s, X3 + 2.70, y + 0.10, FULL - 3.00, h - 0.20,
               [L(a, size=10.4, color="D6E0EE")])
        y += h + gap
    d.text(s, X3, y + 0.04, FULL, 0.34,
           [L(c["foot"], size=10.6, bold=True, color="8FA3BF")])


ORDER = [slide_cover, slide_bills, slide_three, slide_gifts, slide_money, slide_fears,
         slide_people, slide_notdo, slide_need, slide_budget, slide_steps, slide_accept,
         slide_fit, slide_push, slide_next, slide_cheat]


def build(footer, sections):
    """sections: 与 ORDER 等长的内容字典列表。"""
    assert len(sections) == len(ORDER), (len(sections), len(ORDER))
    d = Deck(footer)
    for fn, sec in zip(ORDER, sections):
        fn(d, sec)
    return d
