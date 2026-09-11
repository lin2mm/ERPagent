# 打开方式 · 下载指引

---

## 本次新增 ①：目标企业名单 Excel 版（可直接筛选 / 排序 / 打印）

```text
prospects/Target-Companies.xlsx
```

一张工作簿、三张表：

| 工作表 | 里面是什么 |
| --- | --- |
| **名单** | 28 家企业 · 14 列：优先级 / 企业名称 / 省·市 / 网址 / 行业细分 / 主营业务 / 适配案例 / 初筛分 / 匹配点（成功面）/ 风险点（失败面·待核实）/ 建议切入点 / 接触路径 / 可靠度。已冻结表头、开好筛选、按 A4 横向排好版，可直接打印 |
| **评分口径** | 企业侧 E1–E7（35 分）与老板侧 B1–B7（35 分）的打分标准、S/A/B/C 定级线、两条修正规则，以及「初筛分为什么不满分」 |
| **使用说明与获客渠道** | 三步用法（10 分钟桌面核实 → 3 个探路电话 → 五步阶梯）+ 四种持续获客渠道（招标中标名单 / 产业带集群 / 展会 / 招聘信号）+ 免责与保密说明 |

CSV 原样保留（`prospects/Target-Companies.csv`，脚本与系统的通用格式）；Excel 由它生成，
表里的数字与 CSV 完全一致。

---

## 本次新增 ②：老板版简报插入「系统实景」（原文档的图，已去掉公司名）

您要的：**把原文档里的系统截图插进之前那两份老板版 PPT / PDF，但图里的公司名要去掉。**

- **案例 B（非标定制）老板版**：原方案文档里有 5 张架构图，已全部插入，新增第 **05 / 06 / 07** 三页：

  | 新页 | 插了什么 | 一句话说明 |
  | --- | --- | --- |
  | 05 | 总体架构（五层） | 八个入口 → 一条业务流 → 能力中台 → 车间边缘节点 → 设备与 IoT |
  | 06 | 一码到底 + 两段拆单 | 一个解决「查得到」，一个解决「交期 40 → 25 天」 |
  | 07 | 角色 × 菜单权限 + 八周冲刺 | 一个回答"谁会看到什么"，一个回答"多久看到东西" |

  这一份因此从 16 页变成 **19 页**（封面 + 正文 + 速查，页码自动重排）。

- **公司名怎么去的**：原图标题写的是"××门业智能制造一体化系统 · 总体架构"，
  现在用**同色同字号**替换为"非标门智能制造一体化系统 · 总体架构" —— 不是打码、
  不是马赛克，版面看不出动过；**图里的流程、角色、菜单、权限矩阵、说明文字全部保留**。
  处理过程写在 `tools/redact_source_images.py` 里，可复查、可复现。

- **案例 A（卷材 MES）老板版**：原始手册（132 页操作手册）里有约 **99 张真实的系统界面截图**，
  但当初只保留了手册里的**文字**，PDF 本身不在本地 —— 所以这一份暂时还是 16 页。
  **请把那份手册 PDF 再发我一次**（拖进对话即可），我按同样的版式把界面截图插进去，
  并同样处理掉企业标识。版式位置已经留好，拿到文件当天就能出。

- 本次没有动：机制版（18 页 ×2）、8 页 20 分钟面谈版 ×2（这两份刻意保持轻，不放图）。

---

## 两份「老板决策简报」PDF（售前说服版）

用来**提高成交率**的两份 PDF。它们和原来的方案讲的是同一件事，
但换了讲法：原来讲"系统长什么样"，现在讲"老板为什么要买、您能得到什么、要出什么"。

```
pdf/PITCH-01-Roll-Material-MES-Boss-Briefing.pdf     16 页 · 3.5 MB · 卷材 MES
pdf/PITCH-02-Custom-Door-Boss-Briefing.pdf           19 页 · 4.3 MB · 非标定制（含 3 页系统实景）
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

## 本次新增：两份「20 分钟面谈版」PDF（8 页）

同一套内容，只留最关键的 8 页，适合**约 20 分钟的面谈**（不带封面，翻开第一页就进正题）：

```
pdf/PITCH-01-Short-Roll-Material-MES.pdf     8 页 · 1.8 MB · 卷材 MES
pdf/PITCH-02-Short-Custom-Door.pdf           8 页 · 1.8 MB · 非标定制
```

保留的 8 页（就是 16 页版里的第 02/03/04/05/11/13/15/16 页）：

```
① 先说账：您每年最贵的四样东西
② 说到底您买的是三件事
③ 这套系统给您的八样好处
④ 算一笔账（算法与口径，不是承诺值）
⑤ 多久见效：五步推进，每步都能停
⑥ 什么样的厂适合做、什么不适合
⑦ 下一步：第一件事怎么开始
⑧ 一页速查（可拿给股东看）
```

**和 16 页版的分工**：16 页版是**留底材料**，客户会后自己看；8 页版是**面谈用**，
一套讲完 20 分钟，现场就能把话说到"下一步做什么"。
两个版本内容完全一致，不会出现两份材料说法不一样的情况。

可编辑版本（PPTX）：`ppt/PITCH-01-Short-Roll-Material-MES.pptx`、`ppt/PITCH-02-Short-Custom-Door.pptx`。
逐页图片：`download/PITCH-01-Short-slides-png/`、`download/PITCH-02-Short-slides-png/`，各 8 张。

---

## 本次新增：目标客户筛选清单（A4 打印版）

一份**可打印、可填写**的售前筛选打分表，用来决定在一家客户身上投多少售前资源：

```
pdf/CHECKLIST-Customer-Screening.pdf     A4 两页 · 720 KB · 直接打印
templates/07-Customer-Qualification-Checklist.md    可编辑版（Markdown）
```

七步走完一张表：

```
第 0 步  一票否决（6 条，任一不过直接放弃）
第 1 步  企业侧打分（7 项 × 0–5 分 = 35 分）
第 2 步  老板侧打分（7 项 × 0–5 分 = 35 分）
第 3 步  负面清单（10 条，带 ★ 的直接放弃）
第 4 步  定级与打法（S / A / B / C 四级 + 两条修正规则）
第 5 步  见面必问的 12 个问题
第 6 步  走车间 30 分钟看 8 个信号
第 7 步  结论记录（当场写下定级与下一步）
```

**用法**：每接触一家企业填一份，10 分钟填完。核心是两条 ——
企业侧与老板侧**同权**（企业好但老板不动，项目必败），以及
**B1「老板愿不愿意亲自下场」是最强的单一预测因子**。

---

## 本次新增：目标企业名单（28 家初筛）

按 `templates/07` 的打分口径对公开信息做初筛，得出的**线索池**。
每家企业都写了**成功面（为什么值得谈）**与**失败面（要核实什么）**：

```
docs/08-Prospect-Shortlist.md            完整名单 + 评分速览 + 详解 + 获客渠道
prospects/Target-Companies.csv           28 家表格（Excel / Numbers 可直接打开）
```

构成：S 级 4 家 · A 级 11 家 · B 级 11 家 · C 级 2 家，覆盖
卷材加工（案例 A）、非标门/防火门、医疗与实验室家具、展示道具、集成房屋（案例 B）。

**核心校准**：判断规模**按员工数（80–800 人）看，不要按营收** ——
材料与工程类企业的营收被原材料和工程款放大得很厉害。

文档里另附四种**持续获客渠道**：招标中标名单、产业带集群、行业展会、招聘信息
（在招 MES / 信息化岗位 = 已有预算，优先级直接提升）。

> ⚠️ 名单为公开信息初筛，不是尽调结论；本仓库为公开仓库，如需保密请移入私有仓库。

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
pdf/PITCH-01-Short-Roll-Material-MES.pdf             8 页   ← 面谈用（20 分钟）
pdf/PITCH-02-Short-Custom-Door.pdf                   8 页   ← 面谈用（20 分钟）
pdf/PITCH-01-Roll-Material-MES-Boss-Briefing.pdf    16 页   ← 留底用（完整版）
pdf/PITCH-02-Custom-Door-Boss-Briefing.pdf          16 页   ← 留底用（完整版）
pdf/PPT-01.pdf                                      18 页   ← 机制版
pdf/PPT-02.pdf                                      18 页   ← 机制版
```

双击即可用 Keynote / 预览 / Chrome 打开。**不会出现任何格式问题。**
（对 Keynote 来说，把 PDF 拖进去就能当幻灯片用：Keynote → 文件 → 导入。）

### ② PPTX（可编辑，已修复兼容性）

```
ppt/PITCH-01-Short-Roll-Material-MES.pptx             8 页
ppt/PITCH-02-Short-Custom-Door.pptx                   8 页
ppt/PITCH-01-Roll-Material-MES-Boss-Briefing.pptx    16 页
ppt/PITCH-02-Custom-Door-Boss-Briefing.pptx          16 页
ppt/PPT-01-Roll-Material-MES-System-Proposal.pptx    18 页
ppt/PPT-02-Custom-Door-Smart-Manufacturing-Plan.pptx 18 页
```

用 Keynote 打开：**文件 → 打开 → 选择 pptx**。
如果 Keynote 仍提示导入选项，选「导入为幻灯片」即可。

> 若打开后字体与预览图略有差异（Keynote 对「微软雅黑」的处理），
> 属于字体替换，不影响内容；想要 100% 一致的观感请用 PDF。

### ③ 逐页 PNG（绝对保险）

```
download/PITCH-01-Short-slides-png/    8 张
download/PITCH-02-Short-slides-png/    8 张
download/PITCH-01-slides-png/         16 张
download/PITCH-02-slides-png/         16 张
download/PPT-01-slides-png/           18 张
download/PPT-02-slides-png/           18 张
```

每页一张 200 dpi 的图片，任何设备都能打开，适合发微信 / 贴进文档。

---

## 全部材料打包

```
download/ALL-MATERIALS.zip     共 136 个文件，约 42 MB
```

含：6 份 PPTX + 7 份 PDF + 6 张预览总图 + 87 张逐页 PNG + 5 张已脱敏插图
+ 9 份文档 + 7 份模板 + 3 份规格 + 2 份评分报告 + 名单 CSV / Excel + README + 本下载指引。

---

## 直链（GitHub，公开仓库不需登录）

```
# 20 分钟面谈版 PDF（现场用这两份）
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/PITCH-01-Short-Roll-Material-MES.pdf
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/PITCH-02-Short-Custom-Door.pdf

# 售前说服版 PDF（留底、会后发给客户）
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

# 目标客户筛选清单（A4 打印版）
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/CHECKLIST-Customer-Screening.pdf

# 目标企业名单（28 家初筛）
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/prospects/Target-Companies.xlsx
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/prospects/Target-Companies.csv
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/docs/08-Prospect-Shortlist.md

# 全部材料
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/download/ALL-MATERIALS.zip
```

---

## 附：关于文件名的历史改动

仓库内文件名曾含中文，平台文件面板会显示 `Unknown file`，
现已全部改为英文（**文档内容仍是中文**，未作任何删改）。
