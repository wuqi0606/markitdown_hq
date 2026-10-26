# -*- coding: utf-8 -*-
"""补齐 Markdown 中丢失的标题层级。

背景：MarkItDown 靠 Word 的「标题1/2/3」样式映射生成 `#`。
企业制度文件大量章节标题用的是**普通加粗段落**而非标题样式，
转换后层级完全丢失，父子分段的 `\n# ` / `\n## ` 分隔符随之失效。

本模块从正文内容识别伪标题并补层级，规则与
`tools/restore_headings.py`（本地批处理用）保持一致。

三条硬规则：
1. **只有加粗的编号行才算伪标题**。Word 自动编号列表项
   （`1. 发起申请`，非加粗）不是标题，补了会把层级压平。
2. **补出的层级必须深于前面最近的真实标题**，否则出现「H1 套在 H2 里」的倒挂。
3. **目录区块整体跳过**：目录标题与目录条目都不是正文结构。
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

RE_EXISTING_H = re.compile(r"^\s*(#{1,6})\s")
# 空标题：Word 里空段落套了标题样式，markitdown 转成一行光秃秃的 `#`
RE_EMPTY_HEADING = re.compile(r"^\s*#{1,6}\s*$")
RE_FM_DELIM = re.compile(r"^\s*---\s*$")
# 目录标题：`目录` / `目 录` / `## 目录` / `目录：` 都算
RE_TOC_MARK = re.compile(r"^\s*(?:#{1,6}\s*)?目\s*录\s*[:：]?\s*$")
# 目录条目尾部带页码，不是真标题
RE_TOC_TAIL = re.compile(r".*\s\d{1,3}$")
# 标题后面允许跟一小段括号注释
RE_ANNOTATION = re.compile(r"^\s*[（(][^)）]{0,60}[)）]\s*$")

# 伪标题必须**整体加粗**，可选带数字编号前缀，允许尾部跟补充说明：
#   `1. **预算制定**`                    → 前缀 "1. "，标题 "预算制定"
#   `**（二）区域事业部**（指签署…）`     → 标题 "（二）区域事业部"，尾部 "（指签署…）"
RE_BOLD_TITLE = re.compile(
    r"^\s*(?P<pre>(?:\d{1,2}\s*[、.．]\s*)?)\*\*(?P<title>.+?)\*\*(?P<post>.*)$"
)

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


def _bold_parts(raw: str) -> tuple[str, str, str] | None:
    """拆解「加粗标题行」。

    返回 (编号前缀, 加粗标题文本, 加粗之后的补充说明)；不是加粗行返回 None。

    - `1. **预算制定**`                → ("1. ", "预算制定", "")
    - `**（一）定义**`                  → ("", "（一）定义", "")
    - `**（二）区域事业部**（指签署…）`  → ("", "（二）区域事业部", "（指签署…）")
    """
    m = RE_BOLD_TITLE.match(raw)
    if not m:
        return None
    pre = m.group("pre").strip()
    title = m.group("title").strip()
    post = m.group("post").strip()
    if not title:
        return None
    return pre, title, post


def _form_of(text: str) -> str | None:
    """返回标题文本命中的编号形态 key；不是候选标题返回 None。

    `text` 是去掉加粗标记、并且已经剥掉尾部补充说明后的**纯标题文本**。
    """
    if not text or len(text) > MAX_TITLE_LEN:
        return None
    if BAD_CHARS & set(text):
        return None
    if RE_TOC_TAIL.match(text):
        return None
    for key, pat in FORMS:
        if pat.match(text):
            return key
    return None


def _is_annotation(post: str) -> bool:
    """标题后面跟的内容，是不是「补充说明」而不是正文。

    这是区分真标题与「列表项 + 行内加粗」的关键：

    - `**（二）区域事业部**（指签署合同部门…）` → 括号注释，是标题 ✓
    - `2.**当月**线上归档的合同/协议，…`          → 行内加粗强调，不是标题 ✗
    """
    if not post:
        return True
    if "**" in post or "__" in post:
        return False
    if set(post) & set("。；！？"):
        return False
    return bool(RE_ANNOTATION.match(post))


def _classify(raw: str) -> tuple[str, str, str] | None:
    """判定该行是不是「加粗的编号标题」——真伪标题。

    返回 (形态 key, 标题正文, 标题后的补充说明)；不是则返回 None。

    这是区分伪标题与 Word 自动编号列表项的唯一可靠信号：
    - `1. **预算制定**`  → Word 里的加粗段落，是真章节标题，应补 `#`
    - `1. 发起申请`     → Word 的自动编号列表项（numPr），是列表内容，不该补 `#`

    后者若被误升为标题，层级会被压平（列表项与真标题同级），
    父子分段时整节内容会被并进同一个父段。
    """
    parts = _bold_parts(raw)
    if parts is None:
        return None
    pre, title, post = parts
    body = (pre + title).strip()
    key = _form_of(body)
    if key is None:
        return None
    if not _is_annotation(post):
        return None
    return key, body, post


def _real_heading_level(raw: str) -> int | None:
    """该行是否已经是 Markdown 标题；是则返回其层级（1~6）。"""
    m = RE_EXISTING_H.match(raw)
    if not m:
        return None
    return len(m.group(1))


def _normalize_heading(raw: str) -> str:
    """去掉标题行里多余的加粗标记。

    Word 里「标题2 样式 + 加粗」的段落会被 markitdown 转成
    `## **借阅范围与权限**` —— 多余的 `**` 会污染标题文本，
    也会让 `\n# ` 分隔符在带粗体的标题上匹配不到。
    """
    m = RE_EXISTING_H.match(raw)
    if not m:
        return raw
    hashes = m.group(1)
    text = raw[m.end():]
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"__(.+?)__", r"\1", text)
    return f"{hashes} {text.strip()}"


def _build_level_map(lines: list[str]) -> dict[str, int]:
    """扫描全文出现的编号形态，按粗细顺序分配层级。

    例：文档同时出现「一、」和「1、」→ cn=1, num=2（招投标类制度）；
    文档只有「1、」→ num=1。附件始终与中文章节同级。

    只统计**加粗**的编号标题 —— 非加粗的编号行是 Word 自动编号列表项，
    混进来会把层级映射算错。
    """
    present: list[str] = []
    seen = set()
    for line in lines:
        if RE_EXISTING_H.match(line):
            continue
        cls = _classify(line)
        if cls is None:
            continue
        key = cls[0]
        if key not in seen:
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


def restore_headings(text: str) -> tuple[str, int]:
    """补齐标题层级，返回 (处理后文本, 补齐数量)。

    三条硬约束：
    1. 只补**加粗**的编号标题。Word 自动编号列表项（`1. 发起申请`）不是标题，
       补了会把层级压平。
    2. 补出来的层级必须**深于**它前面最近的真实标题（Word 标题样式转出来的
       `#` / `##`）。否则会出现「H1 出现在 H2 里面」的倒挂，
       父子分段时整节内容会被并进同一个父段。
    3. 目录区块整体跳过（目录标题行本身也会被删掉）。
    """
    lines = text.split("\n")
    level_map = _build_level_map(lines)  # 先扫一遍定层级，再逐行改写
    out: list[str] = []
    in_fm = False
    in_toc = False
    added = 0
    seen: set[str] = set()
    real_level = 0  # 前一个真实标题的层级；0 = 还没有真实标题

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

        # 目录标题行：删掉它，并进入目录区块（区块内的条目一并跳过）
        if RE_TOC_MARK.match(line):
            in_toc = True
            continue
        if in_toc and line.startswith("#"):
            in_toc = False

        if in_toc:
            continue

        # 空标题（Word 里空段落套了标题样式）直接丢掉
        if RE_EMPTY_HEADING.match(line):
            continue

        # 真实标题：记录其层级，作为后续伪标题的层级下界
        existing = _real_heading_level(line)
        if existing is not None:
            real_level = existing
            out.append(_normalize_heading(line))
            continue

        # 非加粗的编号行 = Word 自动编号列表项，保持原样
        cls = _classify(line)
        if cls is None:
            out.append(line)
            continue

        key, body, post = cls
        lvl = level_map.get(key)
        # 层级锚定：必须深于前面的真实标题，否则倒挂
        if lvl and real_level:
            lvl = max(lvl, real_level + 1)
        if lvl and body not in seen:
            seen.add(body)          # 文档本身有重复章节名，去重
            out.append("#" * lvl + " " + body)
            added += 1
            if post:
                # 标题后面的补充说明（如「（指签署合同部门…）」）保留为正文
                out.append("")
                out.append(post)
        else:
            out.append(line)

    return re.sub(r"\n{3,}", "\n\n", "\n".join(out)), added
