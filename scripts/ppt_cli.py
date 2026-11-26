#!/usr/bin/env python3
"""
cmic-ppt CLI: 弱模型友好的企业模板化原生 PPTX 生成工具。

核心职责：
  1. templates / inspect : 列出模板、版式、槽位字数上限、3D 图标短名及 JSON 示例；
  2. init                : 初始化项目目录并自动拷入模板底图/Logo/图标；
  3. add / apply-deck    : 按 JSON 自动填槽、CJK 感知换行、**重点套色**、XML 转义、图片拷贝；
  4. build               : 占位符检查 -> ppt-master 像素级溢出门禁 -> 编译结构化母版 .pptx -> 可选 WPS/PPT 渲染 PNG；
  5. extract             : 调用 ppt-master source_to_md.py 从旧 PPTX/DOCX/PDF 提取纯文字与图片素材。
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

# 确保 Windows 终端 UTF-8 输出不因 GBK 编码抛错
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SKILL_ROOT = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = SKILL_ROOT / "templates"
INDEX_FILE = TEMPLATES_DIR / "_index.json"

# 避头标点（不允许单独落在新行行首）
NO_LINE_START_PUNCT = set("，。、；：！？）》】”’,.!?;:)]}%~～")
# 避尾标点（尽量不留在行尾）
NO_LINE_END_PUNCT = set("（《【“‘([")

TOKEN_PATTERN = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")
EMPHASIS_SPLIT_RE = re.compile(r"(\*\*[^*\n]+?\*\*)")
ATOMIC_TOKEN_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9._/%+\-]*|\s+|.")


def die(msg: str, code: int = 1) -> None:
    print(f"[ERROR] {msg}", file=sys.stderr)
    sys.exit(code)


def utf8_env() -> dict[str, str]:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def find_ppt_master() -> Path:
    candidates: list[Path] = []
    # 1. 优先使用 Skill 内嵌的独立精简编译引擎（实现零外部 Skill 依赖、开箱即用）
    candidates.append(SKILL_ROOT / "engine")
    env_dir = os.environ.get("PPT_MASTER_DIR")
    if env_dir:
        candidates.append(Path(env_dir))
    home = Path.home()
    candidates.extend([
        home / ".skilldock" / "skills" / "ppt-master",
        home / ".codex" / "skills" / "ppt-master",
        SKILL_ROOT.parent / "ppt-master",
    ])
    for cand in candidates:
        if (
            (cand / "scripts" / "svg_to_pptx.py").is_file()
            and (cand / "scripts" / "svg_quality_checker.py").is_file()
        ):
            return cand
    die(
        "未找到内嵌编译引擎 engine/scripts/，也未找到外部 ppt-master。"
        "请确认完整拉取了 cmic-ppt 仓库，或通过环境变量 PPT_MASTER_DIR 指定路径。"
    )
    raise AssertionError("unreachable")


def load_index() -> list[dict[str, Any]]:
    if not INDEX_FILE.is_file():
        die(f"模板索引不存在: {INDEX_FILE}")
    data = json.loads(INDEX_FILE.read_text(encoding="utf-8"))
    return data.get("templates", [])


def load_template(template_id: str) -> tuple[Path, dict[str, Any]]:
    entries = load_index()
    for entry in entries:
        if entry.get("id") == template_id:
            tdir = TEMPLATES_DIR / entry["path"]
            tfile = tdir / "template.json"
            if not tfile.is_file():
                die(f"模板配置缺失: {tfile}")
            return tdir, json.loads(tfile.read_text(encoding="utf-8"))
    valid = ", ".join(e.get("id", "") for e in entries)
    die(f"未知模板 ID: '{template_id}'。可用模板: {valid}")
    raise AssertionError("unreachable")


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def strip_emphasis(text: str) -> str:
    return re.sub(r"\*\*([^*]+?)\*\*", r"\1", text)


def char_visual_width(ch: str) -> float:
    if ord(ch) < 128:
        if ch in " ilI1.,:;!'|":
            return 0.36
        if ch in "WM@#%&":
            return 0.72
        return 0.54
    return 1.0


def str_visual_width(s: str) -> float:
    return sum(char_visual_width(c) for c in s)


def parse_emphasis_runs(text: str) -> list[tuple[str, bool]]:
    """将文本拆分为 [(片段, 是否强调)]，保证不出现半截标签。"""
    runs: list[tuple[str, bool]] = []
    for part in EMPHASIS_SPLIT_RE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            runs.append((part[2:-2], True))
        else:
            runs.append((part, False))
    return runs


def merge_runs(runs: list[tuple[str, bool]]) -> list[tuple[str, bool]]:
    merged: list[tuple[str, bool]] = []
    for txt, accent in runs:
        if not txt:
            continue
        if merged and merged[-1][1] == accent:
            merged[-1] = (merged[-1][0] + txt, accent)
        else:
            merged.append((txt, accent))
    return merged


def wrap_styled_lines(
    text: str, chars_per_line: float, max_lines: int, balanced: bool = False
) -> tuple[list[list[tuple[str, bool]]], bool]:
    """CJK + 英文词组感知换行，跨行时保持 **强调** 样式不断裂；居中两行自动均衡断行。"""
    lines: list[list[tuple[str, bool]]] = []
    paragraphs = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    pause_punct = set("，。；：、！？,.;:!?")

    for para in paragraphs:
        if not para:
            lines.append([])
            continue

        # 拆成 (原子词块, accent)
        atoms: list[tuple[str, bool]] = []
        for seg, is_acc in parse_emphasis_runs(para):
            for tok in ATOMIC_TOKEN_RE.findall(seg):
                # 单个超长英文/数字词块若超过整行宽度，则按字符拆
                if len(tok) > 1 and str_visual_width(tok) > chars_per_line:
                    for ch in tok:
                        atoms.append((ch, is_acc))
                else:
                    atoms.append((tok, is_acc))

        total_w = sum(str_visual_width(t) for t, _ in atoms)

        # 居中双行均衡断行：仅当整段略超单行且两行可容纳时，优先在逗号/分句处对半切分
        if (
            balanced
            and max_lines == 2
            and len(paragraphs) == 1
            and total_w > chars_per_line
            and total_w <= chars_per_line * 1.92
            and len(atoms) >= 4
        ):
            prefix_w = [0.0]
            for t, _ in atoms:
                prefix_w.append(prefix_w[-1] + str_visual_width(t))

            best_idx: int | None = None
            best_score = 1e9
            for i in range(1, len(atoms)):
                left_w = prefix_w[i]
                right_w = total_w - left_w
                if left_w > chars_per_line + 0.5 or right_w > chars_per_line + 0.5:
                    continue
                next_tok, next_acc = atoms[i]
                prev_tok, prev_acc = atoms[i - 1]
                if len(next_tok) == 1 and next_tok in NO_LINE_START_PUNCT:
                    continue
                if len(prev_tok) == 1 and prev_tok in NO_LINE_END_PUNCT:
                    continue
                score = abs(left_w - right_w)
                if len(prev_tok) == 1 and prev_tok in pause_punct:
                    score -= 6.0
                if prev_acc and next_acc:
                    # 尽量不把同一段 **强调短语** 拦腰斩断
                    score += 3.0
                if score < best_score:
                    best_score = score
                    best_idx = i

            if best_idx is not None:
                left_atoms = atoms[:best_idx]
                right_atoms = atoms[best_idx:]
                while left_atoms and left_atoms[-1][0].isspace():
                    left_atoms.pop()
                while right_atoms and right_atoms[0][0].isspace():
                    right_atoms.pop(0)
                while right_atoms and right_atoms[-1][0].isspace():
                    right_atoms.pop()
                lines.append(merge_runs(left_atoms))
                if right_atoms:
                    lines.append(merge_runs(right_atoms))
                continue

        cur_runs: list[tuple[str, bool]] = []
        cur_w = 0.0

        for idx, (tok, is_acc) in enumerate(atoms):
            w = str_visual_width(tok)
            is_closing = len(tok) == 1 and tok in NO_LINE_START_PUNCT

            if not cur_runs:
                if tok.isspace():
                    continue
                cur_runs.append((tok, is_acc))
                cur_w = w
                continue

            # 避头标点（如句号、逗号、右引号）绝不单独掉到下一行，允许最多 1.35 汉字宽容差
            limit = chars_per_line + (1.35 if is_closing else 0.0)
            if cur_w + w <= limit:
                cur_runs.append((tok, is_acc))
                cur_w += w
            else:
                # 若当前行最后一个字符是避尾标点（左括号/左书名号/左引号），带到下一行开头
                carry: tuple[str, bool] | None = None
                if (
                    len(cur_runs) > 1
                    and len(cur_runs[-1][0]) == 1
                    and cur_runs[-1][0] in NO_LINE_END_PUNCT
                ):
                    carry = cur_runs.pop()
                while cur_runs and cur_runs[-1][0].isspace():
                    cur_runs.pop()
                lines.append(merge_runs(cur_runs))
                if carry is not None:
                    cur_runs = [carry]
                    cur_w = str_visual_width(carry[0])
                    if not tok.isspace():
                        cur_runs.append((tok, is_acc))
                        cur_w += w
                elif tok.isspace():
                    cur_runs = []
                    cur_w = 0.0
                else:
                    cur_runs = [(tok, is_acc)]
                    cur_w = w

        if cur_runs:
            while cur_runs and cur_runs[-1][0].isspace():
                cur_runs.pop()
            if cur_runs:
                lines.append(merge_runs(cur_runs))

    if not lines:
        lines = [[]]
    exceeded = len(lines) > max_lines
    return lines, exceeded


def render_runs_xml(
    runs: list[tuple[str, bool]], accent_color: str, parent_is_bold: bool = False
) -> str:
    parts: list[str] = []
    weight_attr = "" if parent_is_bold else ' font-weight="700"'
    for txt, is_acc in runs:
        esc = xml_escape(txt)
        if is_acc:
            parts.append(f'<tspan fill="{accent_color}"{weight_attr}>{esc}</tspan>')
        else:
            parts.append(esc)
    return "".join(parts)


def render_multiline_xml(
    text: str,
    parent_x: str,
    chars_per_line: float,
    max_lines: int,
    dy: float,
    accent_color: str,
    is_centered: bool = False,
    parent_is_bold: bool = False,
) -> tuple[str, int, bool]:
    lines, exceeded = wrap_styled_lines(
        text, chars_per_line, max_lines, balanced=is_centered
    )
    if not lines:
        return "", 0, False

    dy_str = f"{dy:g}"
    weight_attr = "" if parent_is_bold else ' font-weight="700"'
    out: list[str] = []
    for i, line_runs in enumerate(lines):
        if i == 0:
            out.append(render_runs_xml(line_runs, accent_color, parent_is_bold))
        else:
            # 若整行全为强调色，直接把 fill 放到外层 tspan 上，避免冗余嵌套
            if len(line_runs) == 1 and line_runs[0][1]:
                esc = xml_escape(line_runs[0][0])
                out.append(
                    f'<tspan x="{parent_x}" dy="{dy_str}" fill="{accent_color}"{weight_attr}>{esc}</tspan>'
                )
            else:
                inner = render_runs_xml(line_runs, accent_color, parent_is_bold)
                out.append(f'<tspan x="{parent_x}" dy="{dy_str}">{inner}</tspan>')
    return "".join(out), len(lines), exceeded


def inspect_enclosing_text(svg_text: str, token_pos: int) -> tuple[str, bool, bool]:
    """向上查找包含该占位符的最近 <text ...> 开始标签，返回 (x, is_centered, is_bold)。"""
    start = svg_text.rfind("<text", 0, token_pos)
    if start == -1:
        return "0", False, False
    end = svg_text.find(">", start)
    if end == -1 or end > token_pos:
        return "0", False, False
    open_tag = svg_text[start : end + 1]
    xm = re.search(r'\bx\s*=\s*"([^"]+)"', open_tag)
    x_val = xm.group(1) if xm else "0"
    is_centered = bool(re.search(r'\btext-anchor\s*=\s*"middle"', open_tag))
    is_bold = bool(re.search(r'\bfont-weight\s*=\s*"(700|bold|800|900)"', open_tag))
    return x_val, is_centered, is_bold


def replace_empty_image_tag(svg_text: str, token: str) -> str:
    """当 image 槽位为空时，把对应 <image href="{{SLOT}}" .../> 替换为柔和提示文字，避免失效 href 阻断导出。"""
    img_re = re.compile(
        r"<image\b[^>]*\bhref\s*=\s*\"" + re.escape(token) + r"\"[^>]*/?>"
    )

    def _repl(match: re.Match[str]) -> str:
        tag = match.group(0)
        xm = re.search(r'\bx\s*=\s*"([0-9.]+)"', tag)
        ym = re.search(r'\by\s*=\s*"([0-9.]+)"', tag)
        wm = re.search(r'\bwidth\s*=\s*"([0-9.]+)"', tag)
        hm = re.search(r'\bheight\s*=\s*"([0-9.]+)"', tag)
        if xm and ym and wm and hm:
            w_val = float(wm.group(1))
            cx = float(xm.group(1)) + w_val / 2.0
            cy = float(ym.group(1)) + float(hm.group(1)) / 2.0 + 4.0
            if w_val <= 56:
                return (
                    f'<text x="{cx:g}" y="{cy:g}" font-size="11" font-weight="700" fill="#174074" '
                    f'text-anchor="middle">APP</text>'
                )
            return (
                f'<text x="{cx:g}" y="{cy:g}" font-size="13" fill="#94A3B8" '
                f'text-anchor="middle">待补位截图</text>'
            )
        return ""

    return img_re.sub(_repl, svg_text)


def resolve_image_asset(
    val: str,
    project_dir: Path,
    search_bases: list[Path],
    slide_name: str,
    slot_name: str,
) -> str:
    images_dir = project_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    # 已是相对 ../images/xxx 格式
    if val.startswith("../images/"):
        fname = val[len("../images/") :]
        if (images_dir / fname).is_file():
            return val
        die(f"{slide_name}.{slot_name} 引用的项目内图片不存在: {images_dir / fname}")

    # 直接写项目 images/ 中已有文件名（必须是非绝对路径，防止 pathlib 覆盖前缀）
    if not Path(val).is_absolute() and (images_dir / val).is_file():
        return f"../images/{val}"

    # 从绝对路径或基准目录查找并复制到项目 images/
    cand_paths = [Path(val)]
    if not Path(val).is_absolute():
        for base in search_bases:
            cand_paths.append(base / val)

    for src in cand_paths:
        if src.is_file():
            dst = images_dir / src.name
            if src.resolve() != dst.resolve():
                shutil.copy2(src, dst)
            return f"../images/{dst.name}"

    die(f"{slide_name}.{slot_name} 找不到图片文件: '{val}'")
    raise AssertionError("unreachable")


def render_slide_svg(
    project_dir: Path,
    template_dir: Path,
    template_cfg: dict[str, Any],
    slide_name: str,
    layout_id: str,
    data: dict[str, Any],
    search_bases: list[Path],
    slide_index: int | None = None,
) -> tuple[Path, list[str]]:
    layouts = template_cfg.get("layouts", {})
    if layout_id not in layouts:
        valid = ", ".join(layouts.keys())
        die(
            f"页面 '{slide_name}' 使用了未知版式 '{layout_id}'。"
            f"模板 '{template_cfg['id']}' 可用版式: {valid}"
        )

    layout_cfg = layouts[layout_id]
    svg_src = template_dir / "layouts" / layout_cfg["file"]
    if not svg_src.is_file():
        die(f"版式 SVG 文件缺失: {svg_src}")

    svg_text = svg_src.read_text(encoding="utf-8")
    slots_cfg: dict[str, dict[str, Any]] = layout_cfg.get("slots", {})
    default_accent = template_cfg.get("accent_color", "#C00000")
    icons_map: dict[str, str] = template_cfg.get("icons", {})
    warnings: list[str] = []

    # 提示未知字段（防止弱模型拼错槽位名却无感知）
    unknown_keys = [k for k in data.keys() if k not in slots_cfg]
    if unknown_keys:
        warnings.append(
            f"{slide_name}: 忽略未定义槽位 {', '.join(unknown_keys)}（请用 inspect 核对槽位名）"
        )

    for slot_name, spec in slots_cfg.items():
        token = "{{" + slot_name + "}}"
        if token not in svg_text:
            continue

        slot_type = spec.get("type", "text")
        accent = spec.get("accent_color", default_accent)

        if slot_name in data and data[slot_name] is not None:
            raw_val = str(data[slot_name])
        elif slot_name == "PAGE_NUM" and slide_index is not None:
            raw_val = str(slide_index)
        elif "default" in spec:
            raw_val = str(spec["default"])
        else:
            raw_val = ""
            if slot_type != "image":
                warnings.append(f"{slide_name}.{slot_name}: 未提供内容，已置空")

        if slot_type == "icon":
            if not raw_val:
                raw_val = str(spec.get("default", ""))
            if raw_val in icons_map:
                href = f"../images/{icons_map[raw_val]}"
            elif raw_val.startswith("../images/"):
                href = raw_val
            elif not Path(raw_val).is_absolute() and (project_dir / "images" / raw_val).is_file():
                href = f"../images/{raw_val}"
            elif Path(raw_val).is_file():
                href = resolve_image_asset(
                    raw_val, project_dir, search_bases, slide_name, slot_name
                )
            else:
                valid_icons = ", ".join(icons_map.keys())
                die(
                    f"{slide_name}.{slot_name} 未知图标名 '{raw_val}'。"
                    f"可用 3D 图标短名: {valid_icons}"
                )
            svg_text = svg_text.replace(token, xml_escape(href))
            continue

        if slot_type == "image":
            if not raw_val.strip():
                warnings.append(f"{slide_name}.{slot_name}: 未提供截图路径，已置为“待补位截图”占位")
                svg_text = replace_empty_image_tag(svg_text, token)
            else:
                href = resolve_image_asset(
                    raw_val.strip(), project_dir, search_bases, slide_name, slot_name
                )
                svg_text = svg_text.replace(token, xml_escape(href))
            continue

        # 文本 / 多行文本字数建议检查
        plain = strip_emphasis(raw_val).replace("\n", "")
        max_chars = spec.get("max_chars")
        if max_chars and len(plain) > max_chars:
            warnings.append(
                f"{slide_name}.{slot_name}: 当前 {len(plain)} 字，超过建议上限 {max_chars} 字"
            )

        if slot_type == "multiline":
            cpl = float(spec.get("chars_per_line", 20))
            max_lines = int(spec.get("max_lines", 2))
            dy = float(spec.get("dy", 22))
            while token in svg_text:
                pos = svg_text.find(token)
                parent_x, is_centered, parent_is_bold = inspect_enclosing_text(svg_text, pos)
                xml_frag, line_cnt, exceeded = render_multiline_xml(
                    raw_val,
                    parent_x,
                    cpl,
                    max_lines,
                    dy,
                    accent,
                    is_centered=is_centered,
                    parent_is_bold=parent_is_bold,
                )
                if exceeded:
                    warnings.append(
                        f"{slide_name}.{slot_name}: 换行后共 {line_cnt} 行，超过推荐 {max_lines} 行（可能触发纵向溢出）"
                    )
                svg_text = svg_text[:pos] + xml_frag + svg_text[pos + len(token) :]
        else:
            # 单行 text：换行符折叠为空格，保留 **重点套色**
            single_text = " ".join(raw_val.splitlines()).strip()
            runs = parse_emphasis_runs(single_text)
            while token in svg_text:
                pos = svg_text.find(token)
                _, _, parent_is_bold = inspect_enclosing_text(svg_text, pos)
                xml_frag = render_runs_xml(runs, accent, parent_is_bold=parent_is_bold)
                svg_text = svg_text[:pos] + xml_frag + svg_text[pos + len(token) :]

    out_dir = project_dir / "svg_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    fname = slide_name if slide_name.endswith(".svg") else f"{slide_name}.svg"
    out_path = out_dir / fname
    out_path.write_text(svg_text, encoding="utf-8")
    return out_path, warnings


def do_init(project_dir: Path, template_id: str) -> None:
    tdir, tcfg = load_template(template_id)
    (project_dir / "svg_output").mkdir(parents=True, exist_ok=True)
    images_dir = project_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    assets_dir = tdir / "assets"
    copied = 0
    if assets_dir.is_dir():
        for asset in sorted(assets_dir.iterdir()):
            if asset.is_file():
                shutil.copy2(asset, images_dir / asset.name)
                copied += 1

    meta_file = project_dir / "project.json"
    meta = {
        "template": template_id,
        "template_name": tcfg.get("name", template_id),
        "updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    meta_file.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] 已初始化项目: {project_dir}")
    print(f"     模板: {template_id} ({tcfg.get('name')}) | 已拷入模板素材 {copied} 个")


def cmd_templates(args: argparse.Namespace) -> None:
    entries = load_index()
    if args.json:
        result = []
        for e in entries:
            _, cfg = load_template(e["id"])
            result.append({
                "id": e["id"],
                "name": e["name"],
                "description": e.get("description", ""),
                "accent_color": cfg.get("accent_color"),
                "icons": list(cfg.get("icons", {}).keys()),
                "layouts": {
                    lid: lcfg.get("description", "")
                    for lid, lcfg in cfg.get("layouts", {}).items()
                },
            })
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    print(f"可用企业 PPT 模板（共 {len(entries)} 套）：\n")
    for e in entries:
        _, cfg = load_template(e["id"])
        print(f"■ {e['id']}  —  {e['name']}")
        print(f"  说明: {e.get('description', '')}")
        print(f"  主强调色 (**关键词** 自动套色): {cfg.get('accent_color')}")
        if cfg.get("icons"):
            print(f"  内置 3D 图标短名: {', '.join(cfg['icons'].keys())}")
        print("  可选版式:")
        for lid, lcfg in cfg.get("layouts", {}).items():
            print(f"    - {lid:<22} : {lcfg.get('description', '')}")
        print()


def cmd_inspect(args: argparse.Namespace) -> None:
    _, cfg = load_template(args.template_id)
    layouts = cfg.get("layouts", {})

    if args.layout_id:
        if args.layout_id not in layouts:
            die(
                f"模板 '{args.template_id}' 中不存在版式 '{args.layout_id}'。"
                f"可选: {', '.join(layouts.keys())}"
            )
        target_layouts = {args.layout_id: layouts[args.layout_id]}
    else:
        target_layouts = layouts

    if args.json:
        payload = {
            "template_id": cfg["id"],
            "name": cfg.get("name"),
            "accent_color": cfg.get("accent_color"),
            "icons": cfg.get("icons", {}),
            "layouts": target_layouts,
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    print(f"模板: {cfg['id']} ({cfg.get('name')}) | 主强调色: {cfg.get('accent_color')}")
    if cfg.get("icons"):
        print(f"可用 3D 图标短名: {', '.join(cfg['icons'].keys())}")
    print()

    for lid, lcfg in target_layouts.items():
        print(f"=== 版式: {lid} ({lcfg.get('description', '')}) ===")
        sample_data: dict[str, str] = {}
        for sname, spec in lcfg.get("slots", {}).items():
            stype = spec.get("type", "text")
            parts = [stype]
            if "max_chars" in spec:
                parts.append(f"<= {spec['max_chars']}字")
            if stype == "multiline":
                parts.append(f"{spec.get('max_lines', 2)}行x{spec.get('chars_per_line', 20)}字")
            if "default" in spec:
                parts.append(f"默认={spec['default']!r}")
            if "hint" in spec:
                parts.append(spec["hint"])
            print(f"  - {sname:<16} : {' | '.join(parts)}")

            if stype == "icon":
                sample_data[sname] = str(spec.get("default", "id_card"))
            elif stype == "image":
                sample_data[sname] = "<图片路径，如 screenshots/step1.png>"
            elif "default" in spec:
                sample_data[sname] = str(spec["default"])
            else:
                sample_data[sname] = f"<{spec.get('hint', sname)}>"

        if args.layout_id:
            print("\n可直接复用的 JSON 槽位骨架:")
            print(json.dumps(sample_data, ensure_ascii=False, indent=2))
        print()


def cmd_init(args: argparse.Namespace) -> None:
    do_init(Path(args.project_dir).resolve(), args.template)


def cmd_add(args: argparse.Namespace) -> None:
    project_dir = Path(args.project_dir).resolve()
    meta_file = project_dir / "project.json"
    if not meta_file.is_file():
        die(f"项目未初始化（找不到 {meta_file}），请先运行 init。")
    meta = json.loads(meta_file.read_text(encoding="utf-8"))
    template_id = meta["template"]
    tdir, tcfg = load_template(template_id)

    data_path = Path(args.data).resolve()
    if not data_path.is_file():
        die(f"找不到数据文件: {data_path}")
    data = json.loads(data_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        die("单页数据 JSON 根节点必须是对象 {...}")

    # 尝试从页名前缀数字推断页码（如 02_pain -> 2）
    m = re.match(r"^0*(\d+)", args.slide_name)
    inferred_idx = int(m.group(1)) if m else None

    out_path, warnings = render_slide_svg(
        project_dir=project_dir,
        template_dir=tdir,
        template_cfg=tcfg,
        slide_name=args.slide_name,
        layout_id=args.layout,
        data=data,
        search_bases=[data_path.parent, project_dir, Path.cwd()],
        slide_index=inferred_idx,
    )
    print(f"[OK] 已生成页面: {out_path}")
    for w in warnings:
        print(f"  [WARN] {w}")


def cmd_apply_deck(args: argparse.Namespace) -> None:
    project_dir = Path(args.project_dir).resolve()
    deck_path = Path(args.deck_json).resolve()
    if not deck_path.is_file():
        die(f"找不到 deck.json: {deck_path}")

    deck = json.loads(deck_path.read_text(encoding="utf-8"))
    if not isinstance(deck, dict) or not isinstance(deck.get("slides"), list):
        die('deck.json 格式无效：需要包含 {"template": "...", "slides": [...]}')

    meta_file = project_dir / "project.json"
    template_id = deck.get("template")
    if not template_id:
        if meta_file.is_file():
            template_id = json.loads(meta_file.read_text(encoding="utf-8")).get("template")
        else:
            die("deck.json 未指定 template，且项目尚未初始化。")

    # 若未初始化或切换了模板，自动初始化并拷入底图/图标
    need_init = not meta_file.is_file()
    if not need_init:
        cur_tpl = json.loads(meta_file.read_text(encoding="utf-8")).get("template")
        if cur_tpl != template_id:
            need_init = True
    if need_init:
        do_init(project_dir, template_id)

    tdir, tcfg = load_template(template_id)

    # 清理旧 svg_output/*.svg，防止缩短页数时残留旧页
    svg_out = project_dir / "svg_output"
    svg_out.mkdir(parents=True, exist_ok=True)
    for old in svg_out.glob("*.svg"):
        old.unlink()

    slides = deck["slides"]
    total_warnings: list[str] = []
    for idx, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            die(f"第 {idx} 页配置必须是对象")
        layout_id = slide.get("layout")
        if not layout_id:
            die(f"第 {idx} 页缺少 layout 字段")
        slide_name = slide.get("name") or f"{idx:02d}_{layout_id}"
        slide_data = slide.get("data") or {}
        if not isinstance(slide_data, dict):
            die(f"页面 '{slide_name}' 的 data 字段必须是对象")

        out_path, warns = render_slide_svg(
            project_dir=project_dir,
            template_dir=tdir,
            template_cfg=tcfg,
            slide_name=slide_name,
            layout_id=layout_id,
            data=slide_data,
            search_bases=[deck_path.parent, project_dir, Path.cwd()],
            slide_index=idx,
        )
        print(f"  [{idx:02d}/{len(slides):02d}] {out_path.name}  (版式: {layout_id})")
        total_warnings.extend(warns)

    print(f"[OK] 批量渲染完成，共 {len(slides)} 页 -> {svg_out}")
    if total_warnings:
        print(f"\n[ADVISORY] 共 {len(total_warnings)} 条文案容量提示：")
        for w in total_warnings:
            print(f"  - {w}")


def render_pptx_to_png_windows(pptx_path: Path, preview_dir: Path) -> list[Path]:
    preview_dir.mkdir(parents=True, exist_ok=True)
    for old in preview_dir.glob("slide_*.png"):
        old.unlink()

    ps_script = f"""
$ErrorActionPreference = 'Stop'
$pptxPath = '{str(pptx_path).replace("'", "''")}'
$outDir = '{str(preview_dir).replace("'", "''")}'
$app = $null
$pres = $null
$engine = ''
$hadExisting = $false
try {{
    try {{
        $app = New-Object -ComObject KWPP.Application
        $engine = 'WPS'
    }} catch {{
        $app = New-Object -ComObject PowerPoint.Application
        $engine = 'PowerPoint'
    }}
    try {{ $app.DisplayAlerts = 1 }} catch {{}}
    try {{ $hadExisting = ($app.Presentations.Count -gt 0) }} catch {{}}
    try {{
        $pres = $app.Presentations.Open($pptxPath, -1, 0, 0)
    }} catch {{
        $pres = $app.Presentations.Open($pptxPath, -1, 0, -1)
    }}
    $count = $pres.Slides.Count
    for ($i = 1; $i -le $count; $i++) {{
        $png = Join-Path $outDir ("slide_{{0:D2}}.png" -f $i)
        $pres.Slides.Item($i).Export($png, "PNG", 1280, 720)
    }}
    Write-Output ("RENDER_OK|" + $engine + "|" + $count)
}} finally {{
    if ($pres -ne $null) {{
        try {{ $pres.Close() }} catch {{}}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    }}
    if ($app -ne $null) {{
        $remain = 0
        try {{ $remain = $app.Presentations.Count }} catch {{}}
        if ((-not $hadExisting) -and ($remain -eq 0)) {{
            try {{ $app.Quit() }} catch {{}}
        }}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }}
}}
"""
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".ps1", encoding="utf-8-sig", delete=False
    ) as tf:
        tf.write(ps_script)
        ps_file = Path(tf.name)

    try:
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps_file)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=60,
            )
        except subprocess.TimeoutExpired:
            print(
                "[WARN] WPS/PowerPoint COM 渲染超时（可能本机 WPS 正处于弹窗或占用状态），已跳过预览图渲染；.pptx 已正常导出。",
                file=sys.stderr,
            )
            return sorted(preview_dir.glob("slide_*.png"))
        if proc.returncode != 0:
            print(
                "[WARN] PPTX 渲染 PNG 失败（.pptx 已正常导出）：\n"
                + (proc.stderr.strip() or proc.stdout.strip()),
                file=sys.stderr,
            )
            return sorted(preview_dir.glob("slide_*.png"))
    finally:
        ps_file.unlink(missing_ok=True)

    return sorted(preview_dir.glob("slide_*.png"))


def summarize_pptx_structure(pptx_path: Path) -> str:
    try:
        from pptx import Presentation  # type: ignore
    except ImportError:
        return "(未安装 python-pptx，跳过母版结构摘要)"
    prs = Presentation(str(pptx_path))
    masters = [m.name for m in prs.slide_masters]
    layouts = [l.name for m in prs.slide_masters for l in m.slide_layouts]
    return (
        f"页数={len(prs.slides)} | 母版数={len(masters)} ({', '.join(masters)}) | "
        f"版式数={len(layouts)}"
    )


def cmd_build(args: argparse.Namespace) -> None:
    project_dir = Path(args.project_dir).resolve()
    svg_dir = project_dir / "svg_output"
    if not svg_dir.is_dir():
        die(f"找不到 svg_output 目录: {svg_dir}")
    svgs = sorted(svg_dir.glob("*.svg"))
    if not svgs:
        die(f"{svg_dir} 中没有任何 .svg 文件，请先运行 apply-deck 或 add。")

    # 1. 未替换占位符门禁
    unreplaced: list[str] = []
    for svg in svgs:
        found = sorted(set(TOKEN_PATTERN.findall(svg.read_text(encoding="utf-8"))))
        if found:
            unreplaced.append(f"{svg.name}: {', '.join('{{' + t + '}}' for t in found)}")
    if unreplaced:
        die("检测到未填充的占位符，请补齐对应槽位后再编译：\n  - " + "\n  - ".join(unreplaced))

    ppt_master = find_ppt_master()
    checker = ppt_master / "scripts" / "svg_quality_checker.py"
    compiler = ppt_master / "scripts" / "svg_to_pptx.py"

    # 2. ppt-master 像素级质量门禁
    if not args.skip_check:
        print(f"[1/3] 正在执行 SVG 像素级溢出与结构质检 ({len(svgs)} 页)...")
        proc = subprocess.run(
            [
                sys.executable,
                str(checker),
                str(project_dir),
                "--quick-generate",
                "--canonical-authoring",
                "--stage",
                "final",
                "--json",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=utf8_env(),
        )

        report_file = project_dir / "validation" / "svg_quality_report.json"
        blocking_issues: list[dict[str, Any]] = []
        introduced_issues: list[dict[str, Any]] = []
        if report_file.is_file():
            rep = json.loads(report_file.read_text(encoding="utf-8"))
            blocking_issues = rep.get("categories", {}).get("blocking", {}).get("issues", [])
            introduced_issues = rep.get("categories", {}).get("introduced", {}).get("issues", [])

        if proc.returncode != 0 or blocking_issues:
            print("[FAIL] 质检未通过，发现阻断性排版/结构问题：", file=sys.stderr)
            if blocking_issues:
                for iss in blocking_issues:
                    print(f"  - {iss.get('file', '?')}: {iss.get('message', '')}", file=sys.stderr)
            else:
                print(proc.stdout, file=sys.stderr)
                print(proc.stderr, file=sys.stderr)
            print(
                "\n[修复建议] 请精简 deck.json 中对应页面的文字长度，重新运行 apply-deck 与 build。",
                file=sys.stderr,
            )
            sys.exit(2)

        print(f"      [OK] 质检通过（0 阻断错误，{len(introduced_issues)} 条非阻断提示）")
        for iss in introduced_issues:
            print(f"      [INFO] {iss.get('file', '?')}: {iss.get('message', '')}")
    else:
        print("[1/3] 已按 --skip-check 跳过 SVG 质检")

    # 3. 编译结构化母版 PPTX
    out_pptx = (
        Path(args.output).resolve()
        if args.output
        else (project_dir / f"{project_dir.name}.pptx")
    )
    out_pptx.parent.mkdir(parents=True, exist_ok=True)
    print(f"[2/3] 正在编译原生可编辑 PPTX -> {out_pptx}")
    proc2 = subprocess.run(
        [
            sys.executable,
            str(compiler),
            str(project_dir),
            "--quick-generate",
            "--image-sizing",
            "display",
            "-o",
            str(out_pptx),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=utf8_env(),
    )
    if proc2.returncode != 0 or not out_pptx.is_file():
        print(proc2.stdout, file=sys.stderr)
        print(proc2.stderr, file=sys.stderr)
        die("PPTX 编译失败。")

    struct_info = summarize_pptx_structure(out_pptx)
    size_kb = out_pptx.stat().st_size / 1024.0
    print(f"      [OK] 导出成功 ({size_kb:.1f} KB) | {struct_info}")

    # 4. 可选渲染每页 PNG 预览图
    if args.render:
        if sys.platform != "win32":
            print("[3/3] [WARN] 当前非 Windows 环境，跳过 WPS/PowerPoint COM 渲染。")
        else:
            preview_dir = project_dir / "preview"
            print(f"[3/3] 正在渲染每页 PNG 预览图 -> {preview_dir}")
            pngs = render_pptx_to_png_windows(out_pptx, preview_dir)
            print(f"      [OK] 已生成 {len(pngs)} 张预览图:")
            for p in pngs:
                print(f"        - {p}")
    else:
        print("[3/3] 未开启 --render，跳过 PNG 预览图渲染。")


def cmd_extract(args: argparse.Namespace) -> None:
    ppt_master = find_ppt_master()
    extractor = ppt_master / "scripts" / "source_to_md.py"
    if not extractor.is_file():
        die(f"找不到素材提取脚本: {extractor}")
    src = Path(args.source_file).resolve()
    if not src.is_file():
        die(f"找不到源文件: {src}")
    out_dir = Path(args.output).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    proc = subprocess.run(
        [sys.executable, str(extractor), str(src), "-o", str(out_dir)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=utf8_env(),
    )
    print(proc.stdout)
    if proc.returncode != 0:
        print(proc.stderr, file=sys.stderr)
        die("素材提取失败。")
    print(f"[OK] 提取完成 -> {out_dir}")


def cmd_check_env(args: argparse.Namespace) -> None:
    import importlib

    print("=== cmic-ppt 运行环境自检 ===")
    py_ver = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    py_ok = sys.version_info >= (3, 10)
    print(f"[{'OK' if py_ok else 'FAIL'}] Python 版本: {py_ver} (要求 >= 3.10)")

    engine_dir = find_ppt_master()
    is_bundled = engine_dir.resolve() == (SKILL_ROOT / "engine").resolve()
    print(f"[OK] 编译引擎: {engine_dir} ({'内置独立引擎' if is_bundled else '外部引擎'})")

    required_mods = [
        ("pptx", "python-pptx"),
        ("lxml", "lxml"),
        ("PIL", "Pillow"),
        ("xlsxwriter", "XlsxWriter"),
        ("yaml", "PyYAML"),
    ]
    optional_mods = [
        ("uharfbuzz", "uharfbuzz（高精度字体测宽，建议安装）"),
        ("pathops", "skia-pathops（布尔路径支持，可选）"),
        ("mammoth", "mammoth（提取 Word .docx 文档，可选）"),
        ("fitz", "PyMuPDF（提取 PDF 文档，可选）"),
        ("openpyxl", "openpyxl（提取 Excel 表格，可选）"),
    ]

    missing_req: list[str] = []
    for mod, pkg in required_mods:
        try:
            importlib.import_module(mod)
            print(f"[OK] 必需依赖: {pkg}")
        except ImportError:
            print(f"[FAIL] 缺少必需依赖: {pkg}")
            missing_req.append(pkg)

    for mod, desc in optional_mods:
        try:
            importlib.import_module(mod)
            print(f"[OK] 增强依赖: {desc}")
        except ImportError:
            print(f"[INFO] 未安装增强依赖: {desc}")

    tpls = load_index()
    print(f"[OK] 已加载模板数: {len(tpls)} ({', '.join(t['id'] for t in tpls)})")

    if not py_ok or missing_req:
        req_file = SKILL_ROOT / "requirements.txt"
        print(f"\n请先运行以下命令补齐依赖：\n  pip install -r \"{req_file}\"")
        sys.exit(1)
    print("\n[OK] 环境完全就绪，可以开始生成原生 PPTX！")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ppt_cli.py",
        description="cmic-ppt：模板驱动的企业原生 PPTX 生成 CLI（兼容轻量模型）",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_chk = sub.add_parser("check-env", help="自检 Python 环境、内嵌引擎与依赖包是否就绪")
    p_chk.set_defaults(func=cmd_check_env)

    p_tpl = sub.add_parser("templates", help="列出所有可用模板、版式和内置 3D 图标")
    p_tpl.add_argument("--json", action="store_true", help="输出 JSON 格式")
    p_tpl.set_defaults(func=cmd_templates)

    p_ins = sub.add_parser("inspect", help="查看模板版式的槽位、字数上限及 JSON 示例")
    p_ins.add_argument("template_id", help="模板 ID，如 cmic-white-report")
    p_ins.add_argument("layout_id", nargs="?", help="可选：只查看指定版式并输出 JSON 骨架")
    p_ins.add_argument("--json", action="store_true", help="输出 JSON 格式")
    p_ins.set_defaults(func=cmd_inspect)

    p_init = sub.add_parser("init", help="初始化项目目录并拷入模板母版素材")
    p_init.add_argument("project_dir", help="项目目录路径")
    p_init.add_argument("--template", required=True, help="模板 ID")
    p_init.set_defaults(func=cmd_init)

    p_add = sub.add_parser("add", help="按单页 JSON 渲染并添加/覆盖 1 页 SVG")
    p_add.add_argument("project_dir", help="项目目录路径")
    p_add.add_argument("slide_name", help="页名，如 01_cover 或 02_pain")
    p_add.add_argument("--layout", required=True, help="版式 ID")
    p_add.add_argument("--data", required=True, help="单页槽位 JSON 文件路径")
    p_add.set_defaults(func=cmd_add)

    p_apply = sub.add_parser("apply-deck", help="按整套 deck.json 批量生成全部页面 SVG")
    p_apply.add_argument("project_dir", help="项目目录路径")
    p_apply.add_argument("deck_json", help="整套演示文稿 JSON 文件路径")
    p_apply.set_defaults(func=cmd_apply_deck)

    p_build = sub.add_parser("build", help="质检 SVG 并编译为带真实母版的原生 .pptx")
    p_build.add_argument("project_dir", help="项目目录路径")
    p_build.add_argument("-o", "--output", help="输出 .pptx 路径（默认写入项目目录）")
    p_build.add_argument("--render", action="store_true", help="Windows 下调用 WPS/PPT 渲染每页 PNG 预览图")
    p_build.add_argument("--skip-check", action="store_true", help="跳过像素级质检（仅调试用）")
    p_build.set_defaults(func=cmd_build)

    p_ext = sub.add_parser("extract", help="从旧 PPTX / DOCX 等文档提取文字与图片素材")
    p_ext.add_argument("source_file", help="源文件路径")
    p_ext.add_argument("-o", "--output", required=True, help="提取输出目录")
    p_ext.set_defaults(func=cmd_extract)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
