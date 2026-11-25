# cmic-ppt

面向企业汇报、技术方案与行业峰会的**模板化原生 PPTX 生成 Skill**。专为同时兼容 **DeepSeek / GLM 等轻量模型** 与 **强模型** 而设计：模型只负责定文字与选版式（输出标准 `deck.json`），所有坐标、自动换行、`**重点套色**`、母版继承、像素级防溢出门禁与 `.pptx` 编译均由 `scripts/ppt_cli.py` 自动完成。

## 内置模板（按模板分治，可插拔扩展）

| 模板 ID | 风格名称 | 适用场景 | 内置版式数 |
|---|---|---|---|
| `cmic-white-report` | 中移白底标准汇报风 | 日常方案、商业计划书、管理层汇报 | 7 套（封面/三栏痛点/左右对照/上总下分架构/四步真机流程/三栏价值/尾页） |
| `cmic-dark-tech-arch` | 深蓝科技架构大屏风 | 技术架构拆解、高层大屏专题演讲 | 5 套（深蓝封面/左70%架构+右24%痛点卡/左右落差对照/四步真机展示/三栏深蓝价值） |
| `cmic-3d-summit` | 3D 蓝玻大会视觉风 | 对外行业峰会、技术交流会、品牌发布（预置 8 枚透明 3D 图标，免生图直接用） | 4 套（四步能力+三底座/五要素横排/三节点成果背书/三阶段演进愿景） |

## 运行环境依赖

- Python 3.10+（推荐 `python` / `py -3`）
- 底层结构化 PPTX 编译与质检引擎：[ppt-master](https://github.com/hugohe3/ppt-master)（MIT 协议，默认自动查找 `~/.skilldock/skills/ppt-master`，或通过环境变量 `PPT_MASTER_DIR` 指定）
- 依赖 Python 包：`python-pptx`、`XlsxWriter`、`skia-pathops`、`uharfbuzz`、`PyYAML`
- 可选视觉预览渲染（`--render`）：Windows 下自动调用 WPS 演示（`KWPP.Application`）或 Microsoft PowerPoint COM 接口

## 4 步快速上手

```powershell
# 1. 查看所有模板与可用版式
python scripts/ppt_cli.py templates

# 2. 查看某模板的精确槽位名与字数上限
python scripts/ppt_cli.py inspect cmic-white-report
python scripts/ppt_cli.py inspect cmic-white-report three_col_pain

# 3. 按 deck.json 批量生成页面（自动初始化、拷素材、换行、**套色**）
python scripts/ppt_cli.py apply-deck work/demo work/demo/deck.json

# 4. 像素级防溢出质检 + 编译带真实 PowerPoint 母版的 .pptx + 渲染预览图
python scripts/ppt_cli.py build work/demo -o work/demo/output.pptx --render
```

## `deck.json` 示例

```json
{
  "template": "cmic-white-report",
  "slides": [
    {
      "name": "01_cover",
      "layout": "cover",
      "data": {
        "TITLE": "基于手机号码 ID 体系的\nToken 运营方案",
        "SUBTITLE": "以号为根 · 一键连接大模型",
        "AUTHOR": "中移互联网公司",
        "DATE": "2026年11月"
      }
    },
    {
      "name": "02_closing",
      "layout": "closing",
      "data": {
        "TITLE": "感谢聆听，恳请指正！",
        "SUBTITLE": "中移互联网 · 能力创新中心",
        "AUTHOR": "2026年11月"
      }
    }
  ]
}
```

## 目录结构

```text
cmic-ppt/
├── SKILL.md                      # 精简 Agent 执行协议（4 步 SOP）
├── README.md                     # 使用与扩展说明
├── scripts/
│   └── ppt_cli.py                # 弱模型友好 CLI 引擎包装层
├── references/
│   └── custom-svg-guide.md       # 强模型自定义单页 SVG 与新增模板规范
└── templates/
    ├── _index.json               # 模板总索引
    ├── cmic-white-report/        # 模板 1：中移白底汇报
    ├── cmic-dark-tech-arch/      # 模板 2：深蓝科技架构
    └── cmic-3d-summit/           # 模板 3：3D 蓝玻大会（含 8 枚透明 3D 图标）
```

## 新增企业风格模板

无需改动任何 Python 代码，只需 3 步：
1. 在 `templates/` 下新建 `<template-id>/` 目录，放入 `template.json`、`assets/`（母版底图/Logo/图标）与 `layouts/*.svg`；
2. 在 `templates/_index.json` 中新增一条模板索引；
3. 执行 `python scripts/ppt_cli.py templates` 即可生效。

## 致谢与协议说明

底层结构化 SVG 质检与原生 PowerPoint Slide Master 编译能力复用开源项目 `hugohe3/ppt-master`（MIT License）。
