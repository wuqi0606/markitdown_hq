# -*- coding: utf-8 -*-
"""补齐 Markdown 中丢失的标题层级。

背景：MarkItDown 靠 Word 的「标题1/2/3」样式映射生成 `#`。
企业制度文件大量章节标题用的是**普通加粗段落**而非标题样式，
转换后层级完全丢失，父子分段的 `\n#` / `\n##` 分隔符随之失效。

本模块从正文内容识别伪标题并补层级，规则与
`tools/restore_headings.py`（本地批处理用）保持一致。
"""

import re

# ---- 标题识别规则 ----
# 四种编号形态按「从粗到细」的固定顺序排列；每份文档实际出现哪些形态，
# 决定各形态映射到几级标题（动态层级，见 _build_level_map）。
RE_CN_CHAPTER = re.compile(r"^\s*(?:第[一二三四五六七八九十]+[章节条]|[一二三四五六七八九十]+\s*[、.．])\s*\S")
RE_NUM = re.compile(r"^\s*\d{1,2}\s*[、.．]\s*\S")
RE_ATTACH = re.compile(r"^\s*附\s*件\s*[一二三四五六七八九十\d]+\s*[、.．]?\s*\S")
RE_CN_PAREN = re.compile(r"^\s*[（(]\s*[一二三四五六七八九十]+\s*[)）]\s*\S")
RE_NUM_PAREN = re.compile(r"^\s*[（(]\s*\d{1,2}\s*[)）]\s*\S")

# 形态 key -> 检测正则。顺序即粗细顺序（cn 最粗，nump 最细）
FORMS = [
    ("attach", RE_ATTACH),
    ("cn", RE_CN_CHAPTER),
    ("num", RE_NUM),
    ("cnp", RE_CN_PAREN),
    ("nump", RE_NUM_PAREN),
]
# 附件与中文章节同级（都是章级），不占独立层级
_SAME_LEVEL_AS_PREV = {"attach"}

RE_EXISTING_H = re.compile(r"^\s*#{1,6}\s")
RE_FM_DELIM = re.compile(r"^\s*---\s*$")
RE_TOC_MARK = re.compile(r"^\s*目录\s*$")
# 目录条目尾部带页码，不是真标题
RE_TOC_TAIL = re.compile(r".*\s\d{1,3}$")

MAX_TITLE_LEN = 45
# 注意：不含「：」—— Word 里「评分标准：」这类冒号结尾短标题是真二级标题
BAD_CHARS = set("。；！？，")


def strip_bold(s: str) -> str:
    """去掉 markdown 加粗/斜体标记。

    必须先剥掉 `**`，否则 `**（一）定义**` 匹配不到二级标题规则 ——
    这是标题层级全丢的直接原因。
    """
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"__(.+?)__", r"\1", s)
    s = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"\1", s)
    return s.strip()


def _form_of(raw: str) -> str | None:
    """返回该行命中的编号形态 key；不是候选标题返回 None。"""
    text = strip_bold(raw)
    if not text or len(text) > MAX_TITLE_LEN:
        return None
    if RE_EXISTING_H.match(raw):
        return None
    if BAD_CHARS & set(text):
        return None
    if RE_TOC_TAIL.match(text):
        return None
    for key, pat in FORMS:
        if pat.match(text):
            return key
    return None


def _build_level_map(lines: list[str]) -> dict[str, int]:
    """扫描全文出现的编号形态，按粗细顺序分配层级。

    例：文档同时出现「一、」和「1、」→ cn=1, num=2（招投标类制度）；
    文档只有「1、」→ num=1。附件始终与中文章节同级。
    """
    present: list[str] = []
    seen = set()
    for line in lines:
        key = _form_of(line)
        if key and key not in seen:
            seen.add(key)
            present.append(key)
    level: dict[str, int] = {}
    n = 1
    for key in present:
        if key in _SAME_LEVEL_AS_PREV and level:
            level[key] = min(level.values())
        else:
            level[key] = min(n, 6)
            n += 1
    return level


def _level_of(raw: str, level_map: dict[str, int]) -> int | None:
    key = _form_of(raw)
    if key is None:
        return None
    return level_map.get(key)


def restore_headings(text: str) -> tuple[str, int]:
    """补齐标题层级，返回 (处理后文本, 补齐数量)。"""
    lines = text.split("\n")
    level_map = _build_level_map(lines)  # 先扫一遍定层级，再逐行改写
    out: list[str] = []
    in_fm = False
    in_toc = False
    added = 0
    seen: set[str] = set()

    for i, line in enumerate(lines):
        if i == 0 and RE_FM_DELIM.match(line):
            in_fm = True
            out.append(line)
            continue
        if in_fm:
            out.append(line)
            if RE_FM_DELIM.match(line) and i > 0:
                in_fm = False
            continue

        if RE_TOC_MARK.match(line):
            in_toc = True
            out.append(line)
            continue
        if in_toc and line.startswith("#"):
            in_toc = False

        if in_toc:
            out.append(line)
            continue

        lvl = _level_of(line, level_map)
        body = strip_bold(line) if lvl else None
        if lvl and body and body not in seen:
            seen.add(body)          # 文档本身有重复章节名，去重
            out.append("#" * lvl + " " + body)
            added += 1
        else:
            out.append(line)

    return re.sub(r"\n{3,}", "\n\n", "\n".join(out)), added