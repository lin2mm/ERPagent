# 打开方式 · How to open

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

现在两份 PPT 的检查结果均为 **✓ 全部通过**。

---

## 三种打开方式，按推荐顺序

### ① PDF（最稳，Keynote 原生支持）

```
pdf/PPT-01.pdf      18 页
pdf/PPT-02.pdf      18 页
```

双击即可用 Keynote / 预览 / Chrome 打开。**不会出现任何格式问题。**
（对 Keynote 来说，把 PDF 拖进去就能当幻灯片用：Keynote → 文件 → 导入。）

### ② PPTX（可编辑，已修复兼容性）

```
ppt/PPT-01-Roll-Material-MES-System-Proposal.pptx
ppt/PPT-02-Custom-Door-Smart-Manufacturing-Plan.pptx
```

用 Keynote 打开：**文件 → 打开 → 选择 pptx**。
如果 Keynote 仍提示导入选项，选「导入为幻灯片」即可。

> 若打开后字体与预览图略有差异（Keynote 对「微软雅黑」的处理），
> 属于字体替换，不影响内容；想要 100% 一致的观感请用 PDF。

### ③ 逐页 PNG（绝对保险）

```
download/PPT-01-slides-png/    18 张
download/PPT-02-slides-png/    18 张
```

每页一张 200 dpi 的图片，任何设备都能打开，适合发微信 / 贴进文档。

---

## 全部材料打包

```
download/ALL-MATERIALS.zip     26 个文件，9.3 MB
```

含：2 份 PPTX + 2 份 PDF + 2 张预览总图 + 36 张逐页 PNG
+ 7 份文档 + 6 份模板 + 3 份规格 + 2 份评分报告。

---

## 直链（GitHub，公开仓库不需登录）

```
# PDF（推荐）
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/PPT-01.pdf
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/pdf/PPT-02.pdf

# PPTX（可编辑）
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/ppt/PPT-01-Roll-Material-MES-System-Proposal.pptx
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/ppt/PPT-02-Custom-Door-Smart-Manufacturing-Plan.pptx

# 全部材料
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/download/ALL-MATERIALS.zip
```

---

## 附：关于文件名的历史改动

仓库内文件名曾含中文，平台文件面板会显示 `Unknown file`，
现已全部改为英文（**文档内容仍是中文**，未作任何删改）。
