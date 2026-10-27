# -*- coding: utf-8 -*-
"""标题层级归一化 —— MinerU / 上游文本分支专用（Word 分支不调用）。

## 为什么需要

Dify 父子分段的两条硬事实（1.10.0 源码 + 实测）：

1. 父分隔符是 `\n# `，**只认一级标题**；
2. 检索时返回给模型的是**父块全文** —— 命中的是子块，但
   `api/core/rag/datasource/retrieval_service.py` 会按 ChildChunk 反查
   父 DocumentSegment，把父块内容塞进结果，子块只挂在 child_chunks 上。

于是「父块粒度 = 模型看到的粒度」。而两条输入分支的标题层级并不一致：

- Word 分支：插件把章节补成了 **H1**（第一章 / 一、目的…）→ 父分隔符命中，
  父块 = 一章（实测 291~1536 字）→ 可用。
- MinerU 分支：上游把章节标成 **H2**，H1 只出现在封面 → 父分隔符一个都匹配不到，
  父块退化成「按长度上限硬切」：
    - 短文档：整篇正文一个父块（实测 3122 字，一次检索吐全文）
    - 长文档：切在句子中间、HTML 表格中间（实测 3984 字的半截表格）
  → 分段不可用。

## 规则

1. **上移**：≥2 级标题整体上移一级（H2→H1、H3→H2…），一级标题不动。
   目的：让「章节」落到父分隔符的切点上。
2. **空壳降级**：标题后紧跟下一个标题（中间无正文）→ 该标题降一级。
   这类标题只是过渡（「1. 引言」后面直接是「1.1. 编写目的」，或上游把文档标题
   重复输出了一次），自身没有内容，不该成为切分点；降级后它并入上一块。

代码围栏（``` / ~~~）内的行不参与判定，也不改写。
"""

from __future__ import annotations

import re

# 标题行：`## 标题`（# 后必须有空白，避免把 `#标签` 当标题）
_RE_HEADING = re.compile(r"^(#{1,6})[ \t]+(.*)$")
# 代码围栏
_RE_FENCE = re.compile(r"^\s*(```|~~~)")

MAX_LEVEL = 6


def _code_mask(lines: list[str]) -> set[int]:
    """返回处于代码围栏内的行号集合。"""
    inside: set[int] = set()
    fence: str | None = None
    for i, line in enumerate(lines):
        m = _RE_FENCE.match(line)
        if fence is None:
            if m:
                fence = m.group(1)
                inside.add(i)
        else:
            inside.add(i)
            if m and m.group(1) == fence:
                fence = None
    return inside


def normalize_heading_levels(text: str) -> tuple[str, dict[str, int]]:
    """把标题层级归一到「章节 = 一级标题」。

    返回 (新文本, 统计)。统计键：
      headings —— 标题总数
      shifted —— 因上移而改变层级的标题数
      demoted —— 因空壳判定而降级的标题数
    """
    lines = text.split("\n")
    in_code = _code_mask(lines)
    head_idx = [
        i for i, line in enumerate(lines)
        if i not in in_code and _RE_HEADING.match(line)
    ]
    stats = {"headings": len(head_idx), "shifted": 0, "demoted": 0}
    if not head_idx:
        return text, stats

    # ---- 1. 上移：≥2 级标题整体上移一级，一级标题不动 ----
    levels: dict[int, int] = {}
    for i in head_idx:
        old = len(_RE_HEADING.match(lines[i]).group(1))  # type: ignore[union-attr]
        new = max(1, old - 1)
        levels[i] = new
        if new != old:
            stats["shifted"] += 1

    # ---- 2. 空壳降级：标题后紧跟下一个标题 → 降一级 ----
    for i in head_idx:
        nxt: int | None = None
        for j in range(i + 1, len(lines)):
            if lines[j].strip():
                nxt = j
                break
        if nxt is None or nxt not in levels:
            continue  # 后面是正文（或到文尾）→ 不是空壳
        new = min(MAX_LEVEL, levels[i] + 1)
        if new != levels[i]:
            levels[i] = new
            stats["demoted"] += 1

    # ---- 3. 写回 ----
    out = list(lines)
    for i, lvl in levels.items():
        title = _RE_HEADING.match(out[i]).group(2).rstrip()  # type: ignore[union-attr]
        out[i] = "#" * lvl + " " + title
    return "\n".join(out), stats
