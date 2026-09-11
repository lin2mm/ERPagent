# 下载说明 · How to download

---

## 最快的办法：GitHub 直链（不依赖本平台 UI）

仓库是公开的，**不需要登录**。下面链接可以直接复制到浏览器打开：

### 两份 PPT（点开即下，各约 110 KB）

```
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/ppt/PPT-01-Roll-Material-MES-System-Proposal.pptx

https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/ppt/PPT-02-Custom-Door-Smart-Manufacturing-Plan.pptx
```

### 全部材料打包（3.1 MB，24 个文件）

```
https://github.com/lin2mm/ERPagent/raw/arena/01a08f24-erpagent/download/ALL-MATERIALS.zip
```

### 想先浏览再下

```
https://github.com/lin2mm/ERPagent/tree/arena/01a08f24-erpagent/ppt
```

进去 → 点文件名 → 页面右上角 **Download raw file**（↓ 图标）。

---

## 在本平台下载

文件在**文件查看器**里打开后，下载按钮在**该面板右上角**。

> 注意区分两个面板：
> - **Diff 面板**（顶部有 `Diff +xxx -x | Checks` 标签）—— 这是查看代码改动的面板，里面**没有文件下载按钮**
> - **文件查看器** —— 显示文件实际内容，右上角才有 **↓ 下载按钮**
>
> 如果当前在 Diff 面板，点右上角 **×** 关掉，再从文件列表点文件名即可。

---

## 关于「Unknown file」

之前文件列表里出现 `Unknown file`，是因为**文件名含中文**，平台的 Diff / 文件面板解析不了路径。
**现已把仓库里所有文件名统一改为英文（文件内容仍是中文）**，不会再出现这个问题。

| 中文原文 | 现在的文件名 |
| --- | --- |
| `docs/00-总览与阅读指南.md` | `docs/00-Overview-and-Reading-Guide.md` |
| `docs/01-案例精读-A-卷材制造执行系统.md` | `docs/01-Case-A-Roll-Material-MES.md` |
| `docs/02-案例精读-B-非标门一体化方案.md` | `docs/02-Case-B-Custom-Door-Integration.md` |
| `docs/03-通用蓝图-从两案例提炼的可复用模式.md` | `docs/03-Blueprint-Reusable-Patterns.md` |
| `docs/04-Agent工作流设计-能力地图与工程规范.md` | `docs/04-Agent-Workflow-Design.md` |
| `docs/05-新企业落地SOP-阶段化交付手册.md` | `docs/05-Delivery-SOP-for-New-Enterprise.md` |
| `docs/06-风险-验收与度量.md` | `docs/06-Risk-Acceptance-Metrics.md` |
| `templates/01`~`06` | `templates/01-Enterprise-Profile.md` … `06-Agent-List-and-Separation.md` |
| `review/01-原方案评分报告.md` | `review/01-Source-Materials-Score.md` |
| `review/02-PPT评分报告.md` | `review/02-PPT-Score.md` |
| `ppt/PPT-01-卷材制造执行系统-系统方案.pptx` | `ppt/PPT-01-Roll-Material-MES-System-Proposal.pptx` |
| `ppt/PPT-02-非标定制品智能制造一体化方案.pptx` | `ppt/PPT-02-Custom-Door-Smart-Manufacturing-Plan.pptx` |

---

## 压缩包里有什么

`ALL-MATERIALS.zip`（3.1 MB）共 24 个文件：

```
ppt/       2 份 PPT + 2 张 18 页预览总图
docs/      7 份分析文档（案例精读 / 通用蓝图 / Agent 设计 / 落地 SOP / 风险验收）
templates/ 6 份可填空模板
spec/      3 份机器可读规格（agents / gates / entities）
review/    2 份评分报告
README.md  总览与使用建议
```

---

## 打开 PPT 的注意事项

- 用 **PowerPoint / WPS / Keynote** 均可直接打开编辑
- 中文字体是**微软雅黑**（Windows / Mac 自带），不会串版
- 每页都带了**备注讲稿**：放映时用「演讲者视图」，或编辑时点「视图 → 备注页」查看
- 16:9 比例（13.33 × 7.5 英寸），投影和屏幕演示都适配
