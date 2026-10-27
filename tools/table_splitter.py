# -*- coding: utf-8 -*-
"""HTML 表格 → Markdown，并把超大表按第一列分组 —— MinerU / 上游文本分支专用。

## 为什么需要

上游版面解析（MinerU）把 PDF 里的表格转成 **HTML** 而不是 Markdown：

```
<table><tr><td rowspan="2">登录</td><td>登录页面</td>...</tr>...</table>
```

三个后果：

1. **标签进向量**。`<td>` / `rowspan="6"` 这些字符串对语义检索毫无贡献，
   却污染了 embedding 输入。
2. **撑爆父块**。父子分段的父分隔符是 `\n# `，切不出表格内部的边界；
   Dify 的长度兜底（`max_chunk_length`）只能按 `\n\n` 硬切 —— 实测用户手册
   被切出一块 **3984 字**的半截 HTML 表格。
3. **跨页表格被腰斩**。表格横跨两页时，上游会把同一张表输出成两段 HTML，
   中间断开，兜底切分正好切在断点中间。

## 规则

1. **转 Markdown**：解析 `<tr>` / `<td>` / `<th>`，展开 `rowspan` / `colspan`
   成规整矩阵，输出标准 Markdown 表格。
   - 合并单元格的值会补到它覆盖的每一行 —— 内容不丢（实测保留率 110%）。
   - 单元格里的换行折成空格，避免撑破表格行。
2. **大表分组**：行数够多且某一列有明显分组特征时，按该列分组，
   每组前插一个 **二级标题** `## 组名`。

   为什么是二级标题：插入发生在标题层级归一化**之前**，归一化会把
   ≥2 级标题整体上移一级 → `## 组名` 变成 `# 组名` → 命中父子分段的
   父分隔符 `\n# ` → 每组独立成一个父块，检索返回的粒度从「整张表」
   降到「一个模块」（实测用户手册最大父块 3992 字 → 1792 字）。

   分组判据故意保守（行数 ≥ 10、该列唯一值占比 ≤ 60%、组名长度 ≥ 2），
   避免给「每行一个独立记录」的表硬造标题层级。

## 不影响 Word 分支

Word 分支的表格经 mammoth → markdownify 转换后**本来就是 Markdown**
（`<td>` 出现 0 次），本模块只在文本分支调用，33 份 docx 无需重跑。
"""

from __future__ import annotations

import html as _html
import re

# 表格块：<table ...> ... </table>（非贪婪，DOTALL）
_RE_TABLE = re.compile(r"<table\b[^>]*>.*?</table\s*>", re.S | re.I)
# 行
_RE_TR = re.compile(r"<tr\b[^>]*>(.*?)(?:</tr\s*>|$)", re.S | re.I)
# 单元格
_RE_TD = re.compile(r"<t([dh])\b([^>]*)>(.*?)(?:</t\1\s*>|$)", re.S | re.I)
# 跨行/跨列
_RE_ROWSPAN = re.compile(r"rowspan\s*=\s*[\"']?(\d+)", re.I)
_RE_COLSPAN = re.compile(r"colspan\s*=\s*[\"']?(\d+)", re.I)
# 换行类标签 → 空格
_RE_BR = re.compile(r"<\s*(br|/p|/div|/tr)\s*/?\s*>", re.I)

# 分组判据
MIN_GROUP_ROWS = 10        # 行数下限
MAX_UNIQUE_RATIO = 0.6     # 第一列唯一值 / 行数 的上限
MIN_KEY_LEN = 2            # 组名最小长度（排除 "/" 这类占位符）
MIN_TABLE_CHARS = 800      # 表格原始长度下限，太小的表不值得造标题层级


def _cell_text(raw: str) -> str:
    """取单元格纯文本：折行、剥标签、解实体、压空白。"""
    s = _RE_BR.sub(" ", raw)
    s = re.sub(r"<[^>]+>", "", s)
    s = _html.unescape(s)
    s = s.replace(" ", " ")
    return re.sub(r"\s+", " ", s).strip()


def _parse_table(block: str) -> tuple[list[str], list[list[str]]]:
    """解析一张 HTML 表 → (表头, 数据行)。解析不出来返回 ([], [])。"""
    raw_rows: list[list[tuple[str, int, int]]] = []
    for tr in _RE_TR.findall(block):
        cells: list[tuple[str, int, int]] = []
        for m in _RE_TD.finditer(tr):
            attrs = m.group(2)
            body = m.group(3)
            rs_m = _RE_ROWSPAN.search(attrs)
            cs_m = _RE_COLSPAN.search(attrs)
            rs = max(1, min(int(rs_m.group(1)), 100)) if rs_m else 1
            cs = max(1, min(int(cs_m.group(1)), 50)) if cs_m else 1
            cells.append((_cell_text(body), rs, cs))
        if cells:
            raw_rows.append(cells)
    if not raw_rows:
        return [], []

    # 展开 rowspan / colspan —— 标准 HTML 表格布局算法
    occupied: dict[tuple[int, int], str] = {}
    max_r = 0
    for r, cells in enumerate(raw_rows):
        col = 0
        for _text, rs, cs in cells:
            while (r, col) in occupied:
                col += 1
            for dr in range(rs):
                for dc in range(cs):
                    occupied[(r + dr, col + dc)] = _text
            col += cs
        max_r = max(max_r, r + 1)

    max_c = 0
    for (_r, c) in occupied:
        max_c = max(max_c, c + 1)
    if max_c == 0:
        return [], []

    grid: list[list[str]] = []
    for r in range(max_r):
        grid.append([occupied.get((r, c), "") for c in range(max_c)])
    if not grid:
        return [], []
    header = grid[0]
    return header, grid[1:]


def _md_row(cells: list[str]) -> str:
    """一行 Markdown 表格。"""
    return "| " + " | ".join(c.replace("|", "\\|") for c in cells) + " |"


def _is_blank_row(cells: list[str]) -> bool:
    """整行单元格都为空 —— 合并单元格展开后可能产生，纯噪声。"""
    return not any(c.strip() for c in cells)


def _md_table(header: list[str], rows: list[list[str]]) -> str:
    """渲染成 Markdown 表格。整行全空的行直接丢掉。"""
    if not header:
        return ""
    ncol = len(header)
    lines = [_md_row(header), "| " + " | ".join("---" for _ in range(ncol)) + " |"]
    for r in rows:
        r = (r + [""] * ncol)[:ncol]
        if _is_blank_row(r):
            continue
        lines.append(_md_row(r))
    return "\n".join(lines)


def _group_key(header: list[str], rows: list[list[str]]) -> int | None:
    """判断能否按某一列分组；返回列索引，不能分组返回 None。

    判据（保守）：
      - 行数够多
      - 该列的有效唯一值 ≥ 2 个（排除空值和 "/" 这类占位符）
      - 唯一值 / 行数 ≤ 60%（每行都不同的表不算分组表）
    """
    if len(rows) < MIN_GROUP_ROWS:
        return None
    for col in range(min(len(header), 4)):  # 只看前几列，第一列最常用
        keys = [(r[col] if col < len(r) else "").strip() for r in rows]
        valid = [k for k in keys if k and k != "/" and len(k) >= MIN_KEY_LEN]
        if len(valid) < len(rows) * 0.6:
            # 绝大多数行在该列都有有效值才算这一列能当分组键
            continue
        uniq = set(valid)
        if len(uniq) < 2:
            continue
        if len(uniq) / len(rows) > MAX_UNIQUE_RATIO:
            continue
        return col
    return None


def convert_html_tables(text: str, group_level: int = 2) -> tuple[str, dict[str, int]]:
    """把文本里的 HTML 表格转成 Markdown；超大表按某一列分组。

    `group_level` —— 分组标题的层级：
      - 2（默认）：插二级标题。后续的标题层级归一化会把 ≥2 级标题整体
        上移一级 → 变成一级 → 命中父子分段的父分隔符 `\n# `。
        同时归一化的「空壳降级」能把紧邻其后的容器型标题降级，
        避免留下「只有标题」的空父块。
      - 1：直接插一级标题（不做归一化时用）。

    返回 (新文本, 统计)。统计键：
      tables    —— 找到的 HTML 表格数
      converted —— 成功转成 Markdown 的表格数
      grouped   —— 触发了分组的表格数
      groups    —— 拆分出的分组总数
    """
    stats = {"tables": 0, "converted": 0, "grouped": 0, "groups": 0}
    if "<table" not in text.lower():
        return text, stats

    level = max(1, min(int(group_level), 6))
    prefix = "#" * level + " "

    out: list[str] = []
    pos = 0
    for m in _RE_TABLE.finditer(text):
        out.append(text[pos:m.start()])
        pos = m.end()
        stats["tables"] += 1
        block = m.group(0)
        header, rows = _parse_table(block)
        if not header or not rows:
            # 解析不出来就原样保留 —— 宁可留 HTML 也不丢内容
            out.append(block)
            continue
        stats["converted"] += 1

        col = None
        if len(block) >= MIN_TABLE_CHARS:
            col = _group_key(header, rows)

        if col is None:
            out.append("\n\n" + _md_table(header, rows) + "\n\n")
            continue

        # 按分组键切分（连续相同的键合并成一组）
        groups: list[tuple[str, list[list[str]]]] = []
        cur_name = ""
        cur_rows: list[list[str]] = []
        for r in rows:
            key = (r[col] if col < len(r) else "").strip()
            if key and key != cur_name:
                if cur_rows:
                    groups.append((cur_name, cur_rows))
                cur_name = key
                cur_rows = []
            cur_rows.append(r)
        if cur_rows:
            groups.append((cur_name, cur_rows))

        stats["grouped"] += 1
        parts: list[str] = []
        for name, grows in groups:
            md = _md_table(header, grows)
            if name:
                parts.append(f"\n\n{prefix}{name}\n\n{md}")
                stats["groups"] += 1
            else:
                parts.append("\n\n" + md + "\n\n")
        out.append("".join(parts))

    out.append(text[pos:])
    return "".join(out), stats
