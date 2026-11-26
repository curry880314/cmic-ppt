# cmic-ppt —— 企业模板化原生 PPT 生成 Skill

面向企业内部汇报、技术架构方案与行业峰会的 **开箱即用 AI PPT 生成 Skill**（已完整内嵌精简版编译与质检引擎，**无需再单独安装其他 Skill**）。

核心设计理念是**“人定内容，脚本管排版”**：AI 只需要帮你提炼大纲、选模板版式并输出一份结构化文字稿（`deck.json`）；所有的坐标对齐、中英文自动换行、重点套色、像素级防爆框检查、原生 PowerPoint 母版编译全部由本地脚本自动完成。因此不仅适用于强模型，在 **WorkBuddy + DeepSeek / GLM** 等轻量模型下同样稳定不跑版。

---

## 一、这个 Skill 有哪些核心功能

1. **旧文档素材自动提取**：支持直接丢入旧 `.pptx`、`.docx`、`.pdf`，自动抽取文字结构与图片素材，免去手工搬运。
2. **先审逐页大纲，再动工排版**：默认先输出《逐页文字与版式确认表》，你改完字、拍板后再一键生成 PPT，避免 AI 闷头做十几页不合心意的废稿。
3. **自动换行 + `**重点套色**`**：写文字时只需要用 `**重点词**` 标注，脚本自动按模板规范套红/套金加粗，并遵循中文避头尾标点规则自动换行。
4. **像素级防溢出质检门禁**：编译前自动校验每一页的真实字宽与容器边界，发现超字、撑爆卡片或漏填占位符时直接给出中文精简提示，由 AI 自动改短重编。
5. **真·原生可编辑 `.pptx`（带真实 Slide Master 母版）**：不是把整页糊成图片！导出的 `.pptx` 用 PowerPoint 或 WPS 打开后，所有标题、正文、卡片、图标、配图均可自由选中、改字、挪位。
6. **自动导出逐页预览图（Windows）**：本机装有 WPS 演示或 Microsoft PowerPoint 时，编译可自动导出每页 PNG 预览图，方便快速验收。

---

## 二、内置 3 套企业模板

| 模板 ID | 风格名称 | 适用场景 | 内置版式与特色 |
|---|---|---|---|
| `cmic-white-report` | 中移白底标准汇报风 | 日常工作汇报、商业计划书、管理层评审 | **7 套版式**：封面、三栏痛点、左右对照、上总下分架构、四步真机流程、三栏价值、尾页 |
| `cmic-dark-tech-arch` | 深蓝科技架构大屏风 | 技术方案拆解、架构专题、大屏汇报 | **6 套版式**：深蓝封面、左 70% 三列痛点架构+右 24% 攻克卡、**生态中枢连接架构全景（旧链 vs 极简对比 + 四能力底座 + 四统一规范）**、左右落差对照、四步真机展示、三栏深蓝价值 |
| `cmic-3d-summit` | 3D 蓝玻大会视觉风 | 对外行业峰会、技术交流会、品牌发布 | **4 套版式**：四步能力+三底座、五要素横排、三节点成果背书、三阶段演进愿景；**预置 8 枚透明 3D 图标，免生图直接用** |

> 模板采用插拔式目录结构，后续随时可以往 `templates/` 里新增企业新风格，无需修改任何 Python 代码。

---

## 三、如何安装（适配任意支持 Skill 的 AI Agent）

支持 WorkBuddy、Claude Code、Codex、Cursor、Antigravity 等所有支持 `SKILL.md` 规范的 AI Agent，**在哪个 Agent 里发指令就会装到对应环境**。

### 前置环境
- Python 3.10 及以上（终端输入 `python --version` 确认）
- （可选）Windows 下已安装 WPS 演示或 Microsoft PowerPoint，可开启自动出预览图

### 方式 A：发给你的 AI Agent 一句话自动安装（推荐）

直接把下面这句话发给你正在使用的 AI Agent：

```text
请帮我从 https://github.com/curry880314/cmic-ppt.git 克隆安装 cmic-ppt Skill 到你当前环境默认的 skills 目录（确保目录内直接包含 SKILL.md），运行该目录下的 pip install -r requirements.txt 安装依赖，然后执行 python scripts/ppt_cli.py check-env 验证安装成功。
```

**后续版本一键同步更新**：当仓库新增模板或升级引擎时，发给你正在使用的 Agent 一句话即可：
```text
请进入我本地已安装的 cmic-ppt Skill 目录执行 git pull 拉取最新版本，如有新增依赖自动补装，并运行 python scripts/ppt_cli.py check-env 验证环境正常。
```

### 方式 B：离线压缩包安装（无需 GitHub 账号与外网）

1. 获取 `cmic-ppt.zip` 压缩包；
2. 解压到你所用 Agent 的全局 skills 目录（确保 `cmic-ppt` 目录下直接能看到 `SKILL.md`，常见路径如 `~/.workbuddy/skills/cmic-ppt`、`~/.claude/skills/cmic-ppt`、`~/.codex/skills/cmic-ppt`、`~/.cursor/skills/cmic-ppt` 等；支持界面导入的客户端可直接导入压缩包）；
3. 进入解压后的 `cmic-ppt` 目录，在终端执行一次依赖安装与自检：
   ```powershell
   pip install -r requirements.txt
   python scripts/ppt_cli.py check-env
   ```
   看到 `[OK] 环境完全就绪` 后，重启或刷新当前 Agent 即可生效。

---

## 四、日常怎么用（直接复制的话术）

安装完成后，在你常用的 AI Agent 对话框里像和同事沟通一样下指令即可，**不需要记任何命令**：

### 场景 1：把现有方案 / 旧 PPT 重做成新版汇报
> 把这份文件（或直接粘贴文字）用 `cmic-ppt` 整理成 8 页汇报 PPT，风格选「中移白底标准汇报风」。**先给我输出《逐页文字与版式确认表》，等我改完字再生成 `.pptx`。**

### 场景 2：从零提想法做技术架构大屏汇报
> 用 `cmic-ppt` 的「深蓝科技架构大屏风」帮我规划一份 6 页技术架构汇报，主题是 XXX，核心亮点是 A、B、C。先出逐页文字大纲和对应版式给我确认。

### 场景 3：修改某几页文字后重新出稿
> 第 3 页太长了，帮我精简成 3 条要点；第 5 页标题改成「XXX」，然后重新生成 `.pptx` 并渲染预览图。

---

## 五、命令行参考（开发者 / 手动调试用）

```powershell
# 1. 环境与内置引擎自检
python scripts/ppt_cli.py check-env

# 2. 查看所有内置模板与版式
python scripts/ppt_cli.py templates

# 3. 查看某个模板的精确槽位名与字数上限
python scripts/ppt_cli.py inspect cmic-white-report
python scripts/ppt_cli.py inspect cmic-white-report three_col_pain

# 4. 从旧 PPTX / DOCX / PDF 提取文字与图片素材
python scripts/ppt_cli.py extract "旧方案.pptx" -o work/source

# 5. 根据 deck.json 一键生成全部页面 SVG（自动初始化、拷素材、换行、**套色**）
python scripts/ppt_cli.py apply-deck work/demo work/demo/deck.json

# 6. 像素级溢出质检 + 编译带真实母版的原生 .pptx + 渲染每页 PNG 预览
python scripts/ppt_cli.py build work/demo -o work/demo/output.pptx --render
```

---

## 六、项目目录结构

```text
cmic-ppt/
├── SKILL.md                  # Agent 标准执行协议（4 步 SOP）
├── README.md                 # 本使用说明书
├── requirements.txt          # Python 依赖清单
├── scripts/
│   └── ppt_cli.py            # 弱模型友好 CLI 引擎（填槽 / 换行 / 套色 / 质检 / 编译 / 提取）
├── engine/                   # 内嵌精简版结构化 PPTX 编译与质检引擎（MIT，零外部 Skill 依赖）
├── references/
│   └── custom-svg-guide.md   # 强模型自定义单页 SVG 与扩展新模板规范
├── templates/                # 可插拔企业模板库（当前内置 3 套）
│   ├── _index.json
│   ├── cmic-white-report/
│   ├── cmic-dark-tech-arch/
│   └── cmic-3d-summit/
└── THIRD_PARTY_LICENSES/     # 第三方开源协议声明
```

---

## 七、致谢与协议

本 Skill 内嵌 `engine/` 目录中的结构化 SVG 质检、文档提取与原生 PowerPoint Slide Master 编译核心复用开源项目 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)（MIT License），完整许可证见 [THIRD_PARTY_LICENSES/ppt-master-LICENSE.txt](THIRD_PARTY_LICENSES/ppt-master-LICENSE.txt)。
