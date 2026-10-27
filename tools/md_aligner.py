# -*- coding: utf-8 -*-
"""把 docx 段落元数据对齐回 Markdown 行。

MarkItDown 的输出是「重排」过的：段落被合并、编号被重新分配、表格被压成
Markdown 表格行。要把 XML 元数据用回 Markdown，必须先建立
「docx 段落 ↔ md 行」的对应关系。

对齐策略（实测对齐率 96%）：
1. 规范化：去掉 markdown 标记与所有空白，只留文字骨架
2. 精确匹配优先；失败时退化为「一方包含另一方」的宽松匹配
3. 单调游标：docx 段落按文档顺序推进，md 行号只增不减 ——
   避免同一段文字在文档里出现多次时错配

表格内的段落不参与对齐：mammoth 会把整个表格压成若干行 Markdown，
段级对应关系已经不存在。
"""

from __future__ import annotations

import re

from tools.docx_meta import DocxMeta, ParaMeta

# markdown 行内标记
_RE_MD_MARK = re.compile(r"(\*\*|__|\*|`|^#{1,6}\s+|^\s*[-+*]\s+|^\s*\d{1,3}[.)]\s+)")
_RE_WS = re.compile(r"[\s　]+")


def normalize(text: str) -> str:
    """去掉 markdown 标记与空白，只留文字骨架。"""
    s = _RE_MD_MARK.sub("", text or "")
    s = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", s)      # 图片
    s = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", s)  # 链接
    s = _RE_WS.sub("", s)
    return s.strip()


def _content_of(line: str) -> str:
    """取 md 行里的文字内容（去掉标题/列表标记）。"""
    s = line.strip()
    s = re.sub(r"^#{1,6}\s+", "", s)
    s = re.sub(r"^[-+*]\s+", "", s)
    s = re.sub(r"^\d{1,3}[.)]\s+", "", s)
    s = s.strip()
    # 表格行整行去掉竖线
    if s.startswith("|"):
        s = s.replace("|", "")
    return s


class Alignment:
    """md 行号 → docx 段落元数据 的映射。"""

    def __init__(self) -> None:
        self.by_line: dict[int, ParaMeta] = {}
        self.stats = {"lines": 0, "matched": 0, "loose": 0,
                      "table_skipped": 0, "toc_skipped": 0}

    def get(self, line_no: int) -> ParaMeta | None:
        return self.by_line.get(line_no)

    @property
    def rate(self) -> float:
        return self.stats["matched"] / self.stats["lines"] if self.stats["lines"] else 0.0


def align(markdown: str, meta: DocxMeta | None) -> Alignment:
    """建立 md 行 → docx 段落元数据的对应关系。

    `meta` 为 None（PDF / 纯文本模式）时返回空对齐。
    """
    alg = Alignment()
    if meta is None or not meta.paras:
        return alg

    lines = markdown.split("\n")
    alg.stats["lines"] = len(lines)

    # md 行索引：规范化文本 → 行号列表（保持升序，便于单调推进）
    index: dict[str, list[int]] = {}
    line_norm: list[str] = []
    for i, raw in enumerate(lines):
        n = normalize(_content_of(raw))
        line_norm.append(n)
        if n:
            index.setdefault(n, []).append(i)

    cursor = 0
    for p in meta.paras:
        if p.in_table:
            alg.stats["table_skipped"] += 1
            continue
        if p.in_toc:
            # 目录条目在 markdown 里被清洗规则删掉了，不参与对齐
            alg.stats["toc_skipped"] += 1
            continue
        key = normalize(p.text)
        if not key:
            continue

        hit = -1
        how = ""
        # 1) 精确匹配
        for ln in index.get(key, ()):
            if ln >= cursor:
                hit, how = ln, "exact"
                break
        if hit < 0 and len(key) >= 6:
            # 2) 宽松匹配：一方包含另一方
            #    长度必须相近（>=70%），否则「编制目的」会命中「一、 编制目的1」
            for ln, n in enumerate(line_norm):
                if ln < cursor or not n:
                    continue
                if key in n:
                    if len(key) >= 0.7 * len(n):
                        hit, how = ln, "loose"
                        break
                elif n in key:
                    if len(n) >= 0.7 * len(key):
                        hit, how = ln, "loose"
                        break
        if hit < 0:
            continue

        alg.by_line[hit] = p
        alg.stats["matched"] += 1
        if how == "loose":
            alg.stats["loose"] += 1
        cursor = hit + 1

    return alg
