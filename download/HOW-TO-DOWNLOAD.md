# 打开方式 · 下载指引

---

## 本次新增：两份「老板决策简报」PDF（售前说服版）

用来**提高成交率**的两份 PDF，各 16 页。它们和原来的方案讲的是同一件事，
但换了讲法：原来讲"系统长什么样"，现在讲"老板为什么要买、您能得到什么、要出什么"。

```
pdf/PITCH-01-Roll-Material-MES-Boss-Briefing.pdf     16 页 · 3.5 MB · 卷材 MES
pdf/PITCH-02-Custom-Door-Boss-Briefing.pdf           16 页 · 3.5 MB · 非标定制
```

**16 页的顺序就是谈话的顺序**（每页都在 PPTX 备注里附了讲法）：

```
01 封面（口径声明）
02 先说账：您每年最贵的四样东西        ← 开场是账单，不是功能清单
03 说到底您买的是三件事               ← 少赔钱 / 多接单 / 不用天天盯
04 这套系统给您的八样好处              ← 每一样都换算成钱、天数或风险
05 算一笔账（算法与口径，不是承诺值）    ← 敢当场算，信任就建立
06 您最担心的四件事，我们直接答         ← 顾虑要主动先说
07 谁会反对、为什么、怎么安排           ← 人的位置变了，才是真阻力
08 这套方案明确不做什么                ← 主动减承诺 = 最强的信任动作
09 需要您出的三样东西                  ← 把客户投入讲在前面
10 钱怎么算：一次列全                  ← 不留"第二只靴子"
11 多久见效：五步推进，每步都能停        ← 把第一次风险做小
12 怎么验收：看真机拦不拦得住           ← 三处负面测试
13 什么样的厂适合做、什么不适合          ← 敢于劝退反而被选中
14 担心"推不动"？我们用机制帮您推        ← 老板最深的顾虑
15 下一步：第一件事怎么开始
16 一页速查（可拿给股东看）
```

设计依据：`docs/07-Boss-Perspective-Sales-Strategy.md`。
**另一份案例只需替换内容，版式完全复用**（`tools/pitch_deck.py`）。

可编辑版本（PPTX）：`ppt/PITCH-01-*.pptx`、`ppt/PITCH-02-*.pptx`，各 92 KB。
逐页图片：`download/PITCH-01-slides-png/`、`download/PITCH-02-slides-png/`，各 16 张。

---

## 遇到「file format is invalid」怎么办

**根因已找到并修复。** 旧版 PPTX 里有 **18 处高度为 0 的形状**（用作分隔线），
PowerPoint 能容错显示，但 **Keynote 会直接判定整个文件格式无效**。

新版已把这些线条改成极细矩形，并在构建流程里加了 `audit_pptx()` 自检，
确保**不再出现**以下任何一项：

| 检查项 | 为什么 Keynote 会拒绝 |
| --- | --- |
| zero-extent 形状（cx=0 或 cy=0） | **最常见的元凶**，Keynote 直接报格式无效 |
| `p:cxnSp` 连接线 | 部分版本导入异常 |
| 空文本 run | 部分版本导入异常 |
| `a:rPr` 子元素顺序错误 | OOXML 校验失败 |

现在**四份** PPT 的检查结果均为 **✓ 全部通过**。

---

## 三种打开方式，按推荐顺序

### ① PDF（最稳，Keynote 原生支持）

```
pdf/PITCH-01-Roll-Material-MES-Boss-Briefing.pdf    16 页   ← 售前说服版
pdf/PITCH-02-Custom-Door-Boss-Briefing.pdf          16 页   ← 售前说服版
pdf/PPT-01.pdf                                      18 页   ← 机制版
pdf/PPT-02.pdf                                      18 页   ← 机制版
```

双击即可用 Keynote / 预览 / Chrome 打开。**不会出现任何格式问题。**
（对 Keynote 来说，把 PDF 拖进去就能当幻灯片用：Keynote → 文件 → 导入。）

### ② PPTX（可编辑，已修复兼容性）

```
ppt/PITCH-01-Roll-Material-MES-Boss-Briefing.pptx
ppt/PITCH-02-Custom-Door-Boss-Briefing.pptx
ppt/PPT-01-Roll-Material-MES-System-Proposal.pptx
ppt/PPT-02-Custom-Door-Smart-Manufacturing-Plan.pptx
```

用 Keynote 打开：**文件 → 打开 → 选择 pptx**。
如果 Keynote 仍提示导入选项，选「导入为幻灯片」即可。

> 若打开后字体与预览图略有差异（Keynote 对「微软雅黑」的处理），
> 属于字体替换，不影响内容；想要 100% 一致的观感请用 PDF。

### ③ 逐页 PNG（绝对保险）

```
download/PITCH-01-slides-png/    16 张
download/PITCH-02-slides-png/    16 张
download/PPT-01-slides-png/      18 张
download/PPT-02-slides-png/      18 张
```

每页一张 200 dpi 的图片，任何设备都能打开，适合发微信 / 贴进文档。

---

## 全部材料打包

```
download/ALL-MATERIALS.zip     共 101 个文件，约 32 MB
```

含：4 份 PPTX + 4 份 PDF + 4 张预览总图 + 68 张逐页 PNG
+ 8 份文档 + 6 份模板 + 3 份规格 + 2 份评分报告 + README + 本下载指引。

---

## 直链（GitHub，公开仓库不需登录）

```
# 售前说服版 PDF（先看这两份）
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/PITCH-01-Roll-Material-MES-Boss-Briefing.pdf
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/PITCH-02-Custom-Door-Boss-Briefing.pdf

# 机制版 PDF
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/PPT-01.pdf
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/PPT-02.pdf

# 可编辑 PPTX（四份）
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/ppt/PITCH-01-Roll-Material-MES-Boss-Briefing.pptx
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/ppt/PITCH-02-Custom-Door-Boss-Briefing.pptx
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/ppt/PPT-01-Roll-Material-MES-System-Proposal.pptx
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/ppt/PPT-02-Custom-Door-Smart-Manufacturing-Plan.pptx

# 全部材料
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/download/ALL-MATERIALS.zip
```

---

## 附：关于文件名的历史改动

仓库内文件名曾含中文，平台文件面板会显示 `Unknown file`，
现已全部改为英文（**文档内容仍是中文**，未作任何删改）。
