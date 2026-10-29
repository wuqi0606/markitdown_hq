# -*- coding: utf-8 -*-
"""PDF 分支的标题层级重排（形态序体系）。

背景
----
上游解析器（MinerU）把章节标成二级标题（``## 一、目的``）。父子分段的父分隔符
是 ``\\n# ``，只认一级标题 —— 差一层，父块会退化成按长度硬切。

v0.0.23 的 ``heading_normalizer`` 做了「≥2 级标题整体上移一级」，把章节拉回一级，
父块粒度对了。但它只看标题标记本身，**看不出标题之间的从属关系**，于是：

* 上游误判的标题（把普通条目标成章节）也会被一起上移 → 独立成一个父块，
  把真正的章节从中间切断；
* 章下的条目仍是普通段落 → 子分隔符 ``\\n## `` 无事可切 → 父子结构退化为 1:1。

本模块按**编号形态**重建层级：
    一、目的      → H1（章）    ← 父块边界
    （一）财务中心 → H2（条）   ← 子块边界
    1、负责……    → H3（款）   ← 款属于条的内容，不单独标记

两套编号体系必须分开处理
----------------------
* **形态序**：``一、`` ``（一）`` ``1、`` ``（1）`` —— 层级由作用域栈深度决定。
* **点分十进制**：``1.`` ``1.1.`` ``1.1.1.`` —— 层级由点数决定。

两者混用会互相干扰：形态序里 ``1、`` 是最细的一层，点分体系里 ``1.`` 是最粗的一层。
所以先判定文档属于哪套：

* **形态序** → 按作用域栈重排（``relevel_headings``）
* **点分十进制** → 按编号点数重建（``rebuild_dotted_headings``）

防误伤
------
1. **长度上限**：超长的行不当作标题候选（长段落不是标题）。
2. **序号连续性**：形态序里 ``num`` 形态（``1、`` ``2、``）只有在同一章内
   **序号连续**时才提升为标题；孤立的 ``1、`` 保持正文。
3. 已有 ``#`` 标记的标题一律保留（不丢上游判定）。
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict

from tools.docx_meta import FORM_RANK, form_from_text
from tools.heading_normalizer import _code_mask

RE_HEAD = re.compile(r"^(#{1,6})[ \t]+(.*)$")

# 点分十进制：1. / 1.1. / 1.1.1.（至少带一个点）
RE_DOTTED = re.compile(r"^(\d+(?:\.\d+)+)\s*[.．、]?\s*(\S.*)$")

# 点分十进制标题的深度（点数）。两种写法都要认：
#   「2.2.1. 我的主页」—— 编号后带点
#   「2.9.1.应签台账」  —— 编号后直接接文字
# 分隔处必须是「点+空白」或「纯空白」，否则「3.5万元」这类标题会被误判成章节。
RE_DOTTED_DEPTH_MULTI = re.compile(r"^(\d+(?:\.\d+)+)(?:\.\s*|\s+)(\S.*)$")
RE_DOTTED_DEPTH_ONE = re.compile(r"^(\d+)\.\s+(\S.*)$")

# 形态序
RE_NUM_SEQ = re.compile(r"^(\d{1,2})\s*[、.．]\s*(\S.*)$")
RE_CN_SEQ = re.compile(r"^([一二三四五六七八九十]{1,3})\s*[、.．]\s*(\S.*)$")
RE_NUM_PAREN = re.compile(r"^[（(]\s*(\d{1,2})\s*[)）]\s*(\S.*)$")
RE_CN_PAREN = re.compile(r"^[（(]\s*([一二三四五六七八九十]{1,3})\s*[)）]\s*(\S.*)$")

# 条目行长度上限：超长的行不当作标题候选
MAX_ITEM_LEN = 240

# 形态序体系里，作为「章」的形态（比它更细的都是它的下级）
CHAPTER_FORMS = frozenset({"cn", "chapter", "article"})

# 作用域边界：章级 + 条级。编号序列在每个作用域内重新开始
SCOPE_FORMS = frozenset({"cn", "chapter", "article", "cnp"})


def _strip_head(line: str) -> tuple[bool, str]:
    """返回 (是否标题行, 标题文本)。"""
    m = RE_HEAD.match(line)
    if m:
        return True, m.group(2).strip()
    return False, line.strip()


def has_dotted_numbering(titles: list[str]) -> bool:
    """标题里出现 ``1.1.`` 这类点分十进制 -> 属于点分体系。"""
    return any(RE_DOTTED.match(t) for t in titles)


def _seq_form(text: str) -> tuple[str | None, int | None]:
    """形态序编号 -> (形态, 序号)。不是编号返回 (None, None)。"""
    t = text.strip()
    if not t or len(t) > MAX_ITEM_LEN:
        return None, None
    m = RE_NUM_PAREN.match(t)
    if m:
        return "nump", int(m.group(1))
    m = RE_CN_PAREN.match(t)
    if m:
        return "cnp", None
    m = RE_NUM_SEQ.match(t)
    if m:
        return "num", int(m.group(1))
    m = RE_CN_SEQ.match(t)
    if m:
        return "cn", None
    return None, None


def _is_table_row(line: str) -> bool:
    """Markdown 表格行（``| 1 | 20190102 | ... |``）。

    版本记录表里满是 ``| 1 |`` ``| 2 |``，会被编号正则误当成条目 ——
    表格行的编号是「单元格内容」，不是章节编号，必须排除。
    """
    s = line.strip()
    return s.startswith("|") and s.endswith("|")


def _collect_candidates(lines: list[str]) -> list[tuple[str | None, int | None]]:
    """每行的候选形态。

    标题行也必须带序号
    ------------------
    上游（MinerU）把普通条误判成章节时，会写成 ``## 7、开票时限：``。
    这行的形态是 ``num``、序号是 **7**，它必须参与「序号连续性」判定 ——
    否则序列里会缺 7 而出现空洞：``1,2,3,4,5,6,__,8`` →
    6 与 8 被判为不连续 → 8 判定为孤立条目 → 不提升为标题 →
    留在 7 的正文里 → 8 与 7 落进同一个子分段。
    """
    out = []
    for line in lines:
        is_head, title = _strip_head(line)
        if is_head:
            f = form_from_text(title)
            n = None
            if f == "num":
                # 形态序的阿拉伯数字条目：取序号（超长行由 _seq_form 判定为不取）
                _, n = _seq_form(title)
            elif f is None:
                f, n = _seq_form(title)
            out.append((f, n))
        else:
            if _is_table_row(line):
                out.append((None, None))
                continue
            f, n = _seq_form(title)
            out.append((f, n))
    return out


def _promotable_num(lines: list[str], cand: list[tuple[str | None, int | None]]) -> set[int]:
    """哪些 ``num`` 形态的行可以提升为标题。

    规则：同一作用域内序号连续的那一段才提升；孤立的 ``1、`` 保持正文。
    作用域 = 最近的**章级或条级**标题（``cn`` / ``chapter`` / ``article`` / ``cnp``）——
    同一章下的不同 ``（一）（二）`` 各自是独立作用域，编号都从 1 重新开始。
    已有 ``#`` 标记的 ``num`` 标题一律保留（不丢上游判定）。
    """
    # 作用域分组：遇到章级/条级形态就换作用域
    group = 0
    group_of = [0] * len(cand)
    for i, (f, _) in enumerate(cand):
        if f in SCOPE_FORMS:
            group += 1
        group_of[i] = group

    by_group: dict[int, list[tuple[int, int]]] = defaultdict(list)
    for i, (f, n) in enumerate(cand):
        if f == "num" and n is not None:
            by_group[group_of[i]].append((i, n))

    promote: set[int] = set()
    for items in by_group.values():
        run = [items[0]]
        best = list(run)
        for prev, cur in zip(items, items[1:]):
            if cur[1] == prev[1] + 1:
                run.append(cur)
                if len(run) > len(best):
                    best = list(run)
            else:
                run = [cur]
        if len(best) >= 2:
            promote.update(i for i, _ in best)
        # 上游已标成标题的 num 形态保留
        for i, _ in items:
            if RE_HEAD.match(lines[i]):
                promote.add(i)
    return promote


def relevel_headings(text: str) -> tuple[str, dict[str, int]]:
    """按编号形态重排标题层级。

    返回 (新文本, 统计)。统计键：
      system    —— 识别到的体系：``dotted`` / ``seq``
      h1/h2/h3  —— 各层标题数
      promoted  —— 从正文提升为标题的行数
      kept      —— 保持原样的标题行数
    """
    lines = text.split("\n")
    titles = [RE_HEAD.match(l).group(2).strip() for l in lines if RE_HEAD.match(l)]

    if has_dotted_numbering(titles):
        # 点分体系：一级标题已经是章节，原样保留
        return text, {"system": "dotted", "kept": len(titles)}

    cand = _collect_candidates(lines)
    promote = _promotable_num(lines, cand)

    stack: list[str] = []
    out: list[str] = []
    stats: Counter = Counter()

    def level_for(form: str) -> int:
        """作用域栈：形态 -> 层级。"""
        nonlocal stack
        while stack and FORM_RANK.get(stack[-1], 0) > FORM_RANK.get(form, 0):
            stack.pop()
        if stack and stack[-1] == form:
            return len(stack)
        stack.append(form)
        return len(stack)

    for i, line in enumerate(lines):
        is_head, title = _strip_head(line)
        form, _ = cand[i]

        if form is None:
            out.append(line)
            if is_head:
                stats["kept"] += 1
            continue
        if form == "num" and i not in promote:
            out.append(line)
            if is_head:
                stats["kept"] += 1
            continue

        lv = level_for(form)
        out.append("#" * min(lv, 6) + " " + title)
        stats[f"h{lv}"] += 1
        if not is_head:
            stats["promoted"] += 1

    stats["system"] = "seq"
    return "\n".join(out), dict(stats)


def _dotted_depth(title: str) -> int | None:
    """点分十进制标题的深度（点数）。不是点分编号返回 None。

    ``2.2.1. 我的主页`` → 3；``2.9.1.应签台账`` → 3；``1. 引言`` → 1。
    """
    t = title.strip()
    m = RE_DOTTED_DEPTH_MULTI.match(t)
    if m:
        return len(m.group(1).split("."))
    if RE_DOTTED_DEPTH_ONE.match(t):
        return 1
    return None


# 章级形态序标题：「第一章 引言」「第二章 首页」「第三章 个人中心模块」。
#
# 混合编号体系：用户手册常用「章用中文序号、节用点分编号」的写法 ——
#     第一章 引言          ← 形态序章级
#     1.1. 编写目的        ← 点分（属于第一章）
#     第二章 首页          ← 形态序章级
#     2.1 项目经理工作台    ← 点分（属于第二章）
#
# 这种文档里，「第X章」是**顶层章节**，层级与点分体系的 depth1 相同（都是 H1）。
# 少了这个识别，混合文档的章标题会被当成「编号章节内部的普通小标题」降成正文
# （见 rebuild_dotted_headings 的 demote 分支）→ 父分隔符 `\n# ` 匹配不到 →
# 该章被并进上一个父块，检索时返回的是上一章的内容。
RE_SEQ_CHAPTER = re.compile(r"^\s*第[一二三四五六七八九十百]+\s*[章节部]")


def _followed_by_table(lines: list[str], i: int) -> bool:
    """下一行非空内容是不是表格行。

    ``table_splitter`` 拆分超大表时会插一个二级标题当分组标题，它后面紧跟表格。
    这类标题是**表格的边界**，不能降成正文 —— 降了就等于把整张大表并回上一块。
    """
    for j in range(i + 1, len(lines)):
        s = lines[j].strip()
        if not s:
            continue
        return s.startswith("|")
    return False


def _is_shell_heading(lines: list[str], i: int) -> bool:
    """标题后面紧跟另一个标题（中间没有正文）→ 空壳。

    空壳标题自己没有内容 —— 它不该成为切分点，也不该单独占一个父块。
    """
    for j in range(i + 1, len(lines)):
        s = lines[j].strip()
        if not s:
            continue
        return bool(RE_HEAD.match(lines[j]))
    return False


def rebuild_dotted_headings(text: str) -> tuple[str, dict[str, int]]:
    """点分十进制体系：按**编号点数**重建标题层级。

    为什么
    ------
    上游（MinerU）把点分体系的**所有**章节都标成 ``##``，不区分点数。
    ``heading_normalizer`` 的两条规则（≥2 级上移一级 + 空壳降级）叠加之后，
    **层级实际由「标题后面有没有正文」决定，与编号点数毫无关系**（实测 43/43 吻合）：

        有正文的 depth2/3 → H1（父块）
        空壳的 depth2/3   → H2（并入上一个父块）

    于是**章节标题和它自己的内容被拆到两个父块**：检索命中「2.2. 首页」这个子块时，
    Dify 返回的是父块全文 = 上一节的正文 → 模型答非所问。

    本函数把层级依据换成「编号点数」，与形态序体系「按编号形态重排」思路一致。

    规则
    ----
    * depth1（``1.``）与 depth2（``2.3.``）→ H1（父块边界）
    * depth3（``2.3.1.``）→ H2（子块边界）
    * depth4+ → 正文
    * **形态序章级标题**（``第一章 引言`` ``第二章 首页``）→ H1（父块边界）
    * 编号章节**内部**的非编号标题（「操作步骤」「合同卡片」「注意:」…）→ 正文
    * 编号章节**之前**的非编号标题（如超大表拆分插入的分组标题）→ 保持原级别

    混合编号体系
    ------------
    用户手册常把「章」写成中文序号、「节」写成点分编号：

        第一章 引言
        1.1. 编写目的
        第二章 首页
        2.1 项目经理工作台

    两套编号混在一篇文档里。``第一章`` 这类**章级**标题与点分体系的 depth1
    同级（都是顶层章节），必须走同一条 H1 分支；否则它会落进「非编号标题」
    分支被降成正文 → 父分隔符匹配不到 → 整章被并进上一个父块。

    顶层章节（depth1）
    ----------------
    depth1 是顶层章节（``1. 引言`` ``2. 功能说明``），必须自己**开启**一个父块 ——
    降成正文会让它粘在前一节的末尾，检索「1. 引言」时返回的是上一节的内容。

    但顶层章节常常是**空壳**（后面直接跟第一节，没有自己的正文）。空壳标题
    单独占一个父块也不好看 —— 检索命中它只会拿到几个字。所以空壳的 depth1
    与它后面**紧跟的第一个 depth2** 合并成一个父块（这与 Word 分支「父块首行
    是章标题、章的内容跟在后面」的结构一致），depth2 降为 H2 留在同一块内。
    """
    lines = text.split("\n")
    in_code = _code_mask(lines)
    out: list[str] = []
    stats: Counter = Counter()
    in_numbered = False
    shell_top = False      # 上一个编号标题是空壳的顶层章节（depth1）
    seq_chapter_seen = False   # 文档含形态序章标题 → 点分编号是「章内」的层级

    for i, line in enumerate(lines):
        m = RE_HEAD.match(line)
        if not m or i in in_code:
            out.append(line)
            continue
        title = m.group(2).strip()
        depth = _dotted_depth(title)

        if depth is None:
            # 形态序章级标题（第一章/第二章/第三部分）→ 顶层章节，与 depth1 同级。
            # 混合编号文档里它必须自己开启一个父块。
            if RE_SEQ_CHAPTER.match(title):
                out.append("# " + title)
                stats["h1"] += 1
                stats["seq_chapter"] += 1
                in_numbered = True
                seq_chapter_seen = True
                shell_top = False
                continue
            # 非编号标题
            if in_numbered and not _followed_by_table(lines, i):
                out.append(title)       # 编号章节内部 → 降为正文
                stats["demoted"] += 1
            else:
                out.append(line)        # 独立区域 / 表格分组标题 → 保持原级别
                stats["kept"] += 1
            continue

        in_numbered = True
        if shell_top and depth == 2:
            # 空壳顶层章节后的第一个节 → 与它同父块
            out.append("## " + title)
            stats["h2"] += 1
            stats["merged"] += 1
            shell_top = False
            continue

        # 混合编号文档：点分编号是「章内」的层级，整体下移一级。
        #   第一章 引言          → H1（父块边界）
        #   1.1. 编写目的        → H2（章内节）
        #   1.1.1. 子节           → 正文（见下）
        #   第二章 首页          → H1（父块边界）
        #   2.1 项目经理工作台    → H2（章内节）
        #
        # 否则「第二章」和它下面的 2.1 会被父分隔符 `\n# ` 切开，
        # 检索命中 2.1 时返回的是它自己的标题，而不是第二章的正文。
        #
        # depth3 及更深降为正文：混合文档里 depth2 已经是 H2，
        # depth3 是它的子级 —— 而 Dify 父子分段只有 `\n# ` / `\n## ` 两级，
        # 没有第三级子块可用，留着 `### ` 标记只会变成向量噪音。
        if seq_chapter_seen:
            if depth <= 2:
                out.append("## " + title)   # 章内节 → H2（子块边界）
                stats["h2"] += 1
            else:
                out.append(title)           # 章内子节及更深 → 正文
                stats["demoted"] += 1
            continue

        if depth <= 2:
            # depth1（顶层章节）与 depth2（节级）→ 一级标题 = 父块边界
            out.append("# " + title)
            stats["h1"] += 1
        elif depth == 3:
            out.append("## " + title)
            stats["h2"] += 1
        else:
            out.append(title)           # depth4+ → 正文
            stats["demoted"] += 1

        # 顶层章节若自身没有正文（后面直接跟标题）→ 记下来，等下一个节来合并
        shell_top = depth == 1 and _is_shell_heading(lines, i)

    stats["system"] = "dotted"
    return "\n".join(out), dict(stats)


def demote_deep_headings(text: str, max_level: int = 2) -> str:
    """把三级及更深的标题降回正文。

    款级（``（1）``）属于条的内容，不需要独立标记 —— 留着 ``### `` 会在切分后
    残留进向量文本，是纯噪音。
    """
    out = []
    for line in text.split("\n"):
        m = RE_HEAD.match(line)
        if m and len(m.group(1)) > max_level:
            out.append(m.group(2).strip())
        else:
            out.append(line)
    return "\n".join(out)
