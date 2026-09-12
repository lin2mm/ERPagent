#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""组装最终目标企业名单：prospects/Target-Companies.csv（24 列）

    /tmp/v/bin/python tools/build_targets.py

数据来源（三份，各管一段）：
  · prospects/Target-Companies.csv   —— 上一版名单（14 / 22 / 24 列都能吃），提供「原始列」
  · prospects/regions-social.json    —— 原有 28 家的「区域」与「社媒」（含少量触达补充）
  · prospects/delta-prospects.json   —— 本轮新增的长三角 / 珠三角 16 家（字段完整，自带区域与社媒）

新 schema 在原有 22 列基础上插入两列：
  · 第 3 列「区域」—— 长三角 / 珠三角 / 长三角外的重点；用来把重点区域排到最前
  · 第 13 列「社媒（公众号 / 抖音 / 1688 / 其他）」—— 触达路径的线上入口

可反复执行：已经是 24 列时，会先剥掉这两列还原成基础列，再重新拼，
所以同一份数据跑多少次结果都一样，不会叠加。
"""
import csv
import json
import os

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
CSV = os.path.join(ROOT, "prospects", "Target-Companies.csv")
DELTA = os.path.join(ROOT, "prospects", "delta-prospects.json")
REGIONS = os.path.join(ROOT, "prospects", "regions-social.json")

# 原始 14 列（脚本认识的最老版本）
BASE14 = ["优先级", "企业名称", "省市", "城市", "网址", "行业细分", "主营业务", "适配案例",
          "预估企业侧分(满分35)", "匹配点（成功面）", "风险点（失败面·待核实）",
          "建议切入点", "接触路径", "信息可靠度"]

# 22 列版（上一版：含官网核实与联系方式）
BASE22 = ["优先级", "优先级建议", "企业名称", "省市", "城市",
          "官网（本轮核实）", "官网主营（反向验证）", "验证结论与修正",
          "联系电话", "邮箱", "详细地址", "触达路径（细化到联系方式）",
          "行业细分", "主营业务（名单口径）", "适配案例", "初筛分(企业侧/35)",
          "成功面", "失败面·待核实", "建议切入点", "来源", "信息可靠度", "核实日期"]

# 24 列版（本版：+ 区域 + 社媒）
HEAD24 = (BASE22[:2] + ["区域"] + BASE22[2:11]
          + ["社媒（公众号 / 抖音 / 1688 / 其他）"] + BASE22[11:])

REGION_ORDER = {"长三角": 0, "珠三角": 1}
DELTA_FIELDS = ["优先级", "优先级建议", "区域", "企业名称", "省市", "城市", "官网",
                "官网主营", "反向验证", "电话", "邮箱", "地址", "社媒", "触达路径",
                "行业细分", "主营业务", "适配案例", "初筛分", "成功面", "失败面·待核实",
                "建议切入点", "来源", "信息可靠度", "核实日期"]


def load_base(exclude=()):
    """把当前 CSV 还原成 22 列基础行（无论磁盘上是哪一版）。

    exclude 里的企业名会被丢掉——它们是 delta-prospects.json 管的那批，
    每次都由 JSON 重新写入，避免第二次运行时被当成基础行重复追加。
    """
    rows = list(csv.reader(open(CSV, encoding="utf-8-sig")))
    head, data = rows[0], rows[1:]
    n = len(head)
    if n == 22:
        return head, [r for r in data if r[2] not in exclude]
    if n == 24:
        # 剥掉「区域」(2) 与「社媒」(12)；触达路径里的「｜补充：」是上一轮拼上去的，也要剥掉
        keep = [i for i in range(24) if i not in (2, 12)]
        body = []
        for r in data:
            if r[3] in exclude:
                continue
            r = [r[i] for i in keep]
            r[11] = r[11].split("｜补充：")[0]
            body.append(r)
        return [head[i] for i in keep], body
    if n == 14:
        idx = {h: i for i, h in enumerate(head)}
        out = []
        for r in data:
            pick = lambda h: r[idx[h]] if h in idx else ""
            out.append([r[idx["优先级"]], f"维持 {r[idx['优先级']]} —— 旧版数据，未核实",
                        r[idx["企业名称"]], r[idx["省市"]], r[idx["城市"]],
                        pick("网址"), "（未核实）", "（未核实）",
                        "", "", "", pick("接触路径"),
                        r[idx["行业细分"]], r[idx["主营业务"]], r[idx["适配案例"]],
                        pick("预估企业侧分(满分35)"), r[idx["匹配点（成功面）"]],
                        r[idx["风险点（失败面·待核实）"]], r[idx["建议切入点"]],
                        "", r[idx["信息可靠度"]], ""])
        return BASE22, out
    raise SystemExit(f"无法识别的名单列数：{n}（应为 14 / 22 / 24）")


def build():
    extra = json.load(open(REGIONS, encoding="utf-8"))["企业"]
    delta = json.load(open(DELTA, encoding="utf-8"))["企业"]
    head, data = load_base(exclude={d["企业名称"] for d in delta})
    assert head == BASE22, "基础列与预期不一致，请检查 CSV 表头"

    out = [HEAD24]
    for r in data:
        row = dict(zip(BASE22, r))
        name = row["企业名称"]
        info = extra.get(name, {})
        region = info.get("区域", "（未分类）")
        social = info.get("社媒", "（待补）")
        path = row["触达路径（细化到联系方式）"]
        if info.get("补注"):
            path = (path + "｜补充：" + info["补注"]) if path else info["补注"]
        out.append([row["优先级"], row["优先级建议"], region, name, row["省市"], row["城市"],
                    row["官网（本轮核实）"], row["官网主营（反向验证）"], row["验证结论与修正"],
                    row["联系电话"], row["邮箱"], row["详细地址"], social, path,
                    row["行业细分"], row["主营业务（名单口径）"], row["适配案例"],
                    row["初筛分(企业侧/35)"], row["成功面"], row["失败面·待核实"],
                    row["建议切入点"], row["来源"], row["信息可靠度"], row["核实日期"]])

    for d in delta:
        missing = [f for f in DELTA_FIELDS if f not in d]
        if missing:
            raise SystemExit(f"{d.get('企业名称')} 缺字段：{missing}")
        out.append([d[f] for f in DELTA_FIELDS])

    def key(row):
        region = REGION_ORDER.get(row[2], 2)
        try:
            score = -float(row[17])
        except ValueError:
            score = 0
        return (region, score, row[3])

    head_row, body = out[0], sorted(out[1:], key=key)

    with open(CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(head_row)
        w.writerows(body)

    n_yd = sum(1 for r in body if r[2] == "长三角")
    n_zs = sum(1 for r in body if r[2] == "珠三角")
    print(f"最终名单：{len(body)} 家 × {len(head_row)} 列  →  {CSV}")
    print(f"  长三角 {n_yd} 家 · 珠三角 {n_zs} 家 · 其他 {len(body) - n_yd - n_zs} 家")
    sock = sum(1 for r in body if r[12] and not r[12].startswith("（"))
    print(f"  有官网链接 {sum(1 for r in body if r[6].startswith('http'))} 家 · "
          f"有电话 {sum(1 for r in body if r[9].strip())} 家 · "
          f"有社媒入口 {sock} 家")


if __name__ == "__main__":
    build()
