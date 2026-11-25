# 进阶指南：强模型自定义单页 SVG 规范

> 仅当内置版式骨架无法表达某页特殊拓扑（如超复杂多层架构图、自定义矩阵）且由强模型执行时按需读取；常规页面一律走 `deck.json` + `ppt_cli.py apply-deck`。

## 1. 必须遵守的结构化母版契约
为了让自定义页与模板其他页共享同一套 PowerPoint Slide Master / Layout，自定义 SVG 必须复制当前模板 `<template>/layouts/` 中同母版页的：
- 根节点属性：`lang="zh-CN" width="1280" height="720" viewBox="0 0 1280 720" font-family="Microsoft YaHei, PingFang SC, Arial, sans-serif"`
- `data-pptx-master` / `data-pptx-master-name` / `data-pptx-layout` / `data-pptx-layout-name`
- 全部 `data-pptx-layer="master"` 和 `data-pptx-layer="layout"` 元素（`id`、坐标、颜色必须逐字节一致）
- 全部 `data-pptx-placeholder="..."` 槽位容器（`id`、类型、`data-pptx-idx`、`data-pptx-bounds` 必须一致）

## 2. 正文绘图区域
- 所有自定义卡片、箭头、流程图必须放进 `<g id="content-object-slot" data-pptx-placeholder="object" data-pptx-binding="proxy" ...>` 内，且严格落在该槽位 `data-pptx-bounds` 范围内。
- 多行正文使用单个 `<text>` 加 `<tspan x="与父text一致" dy="行高">` 实现；不要用 HTML 实体（`&nbsp;`、`&mdash;`），直接写 Unicode 字符，仅转义 `&amp;` `&lt;` `&gt;`。
- 完成后直接运行 `ppt_cli.py build <项目目录> --render` 通过 `uharfbuzz` 溢出检测与视觉核验。

## 3. 如何新增一套新公司模板
1. 在 `templates/` 下新建目录 `<new-template-id>/`，包含 `template.json`、`assets/`、`layouts/*.svg`；
2. 在 `templates/_index.json` 中登记一条新模板元数据；
3. 运行 `ppt_cli.py templates` 核实可见，即可立刻被所有模型调用，无需修改任何 Python 代码。
