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
from tools.heading_levels import BODY, CHILD, DEEP, PARENT, LevelScheme
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


def relevel_headings(
    text: str, scheme: LevelScheme | None = None
) -> tuple[str, dict[str, int]]:
    """按编号形态重排标题层级（形态序体系）。

    返回 (新文本, 统计)。统计键：
      system    —— 识别到的体系：``dotted`` / ``seq``
      parent    —— 父块边界标题数
      child     —— 子块边界标题数
      deep      —— 更深层标题数（会被 ``demote_deep_headings`` 降回正文）
      promoted  —— 从正文提升为标题的行数
      kept      —— 保持原样的标题行数

    语义深度 → 角色的映射：第 1 层 = PARENT，第 2 层 = CHILD，
    第 3 层及更深 = DEEP（比子块更深，不与任何分隔符对齐）。
    具体输出几级由 ``scheme`` 决定（见 ``heading_levels``）。
    """
    if scheme is None:
        scheme = LevelScheme.default()
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
        if lv == 1:
            marker = scheme.marker(PARENT)
        elif lv == 2:
            marker = scheme.marker(CHILD)
        else:
            marker = scheme.marker(DEEP, lv - 2)
        out.append(marker + title)
        if lv == 1:
            stats["parent"] += 1
        elif lv == 2:
            stats["child"] += 1
        else:
            stats["deep"] += 1
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


# 章级形态序标题：第X章 / 第X部分
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

# 章级标题形态（带捕获组）：第X章 / 第X部分
RE_CHAPTER_TITLE = re.compile(r"^\s*第([一二三四五六七八九十百]+)\s*([章节部])")

# 点分小节的首段数字（用于验证「章号 ↔ 节号前缀」一致）
#   `# 3.1 交付清单` → 3 ；`# 1.1. 编写目的` → 1
RE_DOTTED_FIRST = re.compile(r"^#*\s*(\d+)(?:\.\d)+")

# 中文数字 → 整数（章号用不到一百以上）
_CN_DIGITS = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
              "六": 6, "七": 7, "八": 8, "九": 9}

# 标题正文里出现这些字符 → 是句子，不是标题
_TITLE_SENTENCE_CHARS = set("。；！？，、：")
# 标题正文里出现 ASCII 句点 → 多半是目录残留（「个人中心模块.........17」）
_TITLE_DOT = "."
# 句子常见虚词（用于识别「第三章讲了…」这种正文行）。
# 注意：**不能**用「个」「及」「以」等 —— 真实章节标题「个人中心模块」
# 就含「个」，误伤会漏掉最该恢复的那一行。只保留几乎不出现在
# 章节名里的强虚词。
_TITLE_FUNCTION_WORDS = set("了的着过被把给从就都也还并而")

# 章标题正文的最大长度（超出 → 是正文句子）
MAX_CHAPTER_TITLE_LEN = 24

# 章标题后面允许隔几行正文再遇到标题（正文插叙场景）
_CHAPTER_LEAD_SCAN = 6


def _cn_to_int(s: str) -> int | None:
    """中文数字 → 整数（支持 一 ~ 九十九；章号够用）。"""
    if not s:
        return None
    if "十" in s:
        head, _, tail = s.partition("十")
        if head and head not in _CN_DIGITS:
            return None
        if tail and tail not in _CN_DIGITS:
            return None
        tens = _CN_DIGITS[head] if head else 1
        ones = _CN_DIGITS[tail] if tail else 0
        return tens * 10 + ones
    return _CN_DIGITS.get(s)


def _chapter_ordinal(text: str) -> tuple[int, str] | None:
    """「第X章 标题」→ (章号, 标题正文)；不是章标题返回 None。"""
    m = RE_CHAPTER_TITLE.match(text.strip())
    if not m:
        return None
    n = _cn_to_int(m.group(1))
    if n is None:
        return None
    return n, text.strip()[m.end():].strip()


def _title_like(rest: str) -> bool:
    """章标题的正文部分是否像标题（而非一句话）。"""
    if not rest or len(rest) > MAX_CHAPTER_TITLE_LEN:
        return False
    if _TITLE_SENTENCE_CHARS & set(rest):
        return False
    if _TITLE_DOT in rest:
        return False
    if _TITLE_FUNCTION_WORDS & set(rest):
        return False
    return True


def _next_section_first(lines: list[str], i: int, in_code: set[int]) -> int | None:
    """章标题后面第一个标题的首段数字；不是点分标题返回 None。

    扫描时跳过正文行（章标题与第一个小节之间可能有一段引言），
    但一旦遇到非点分标题就放弃 —— 说明这不是「章 → 节」的对应关系。
    """
    seen = 0
    for j in range(i + 1, min(len(lines), i + 1 + _CHAPTER_LEAD_SCAN)):
        if j in in_code:
            continue
        s = lines[j].strip()
        if not s:
            continue
        seen += 1
        if seen > _CHAPTER_LEAD_SCAN:
            return None
        if s.startswith("#"):
            m = RE_DOTTED_FIRST.match(s)
            return int(m.group(1)) if m else None
        if RE_CHAPTER_TITLE.match(s):
            return None          # 下一行就是另一个章 → 不是「章 → 节」
    return None


def _find_bare_chapters(lines: list[str], in_code: set[int]) -> dict[int, str]:
    """找出上游漏标的「第X章」行 → {行号: 标题全文}。

    上游（MinerU）按版面识别标题，章标题落在页底时可能漏标 ——
    实测「第三章 个人中心模块」没带标记，被当成正文，父分隔符匹配不到，
    整章被并进上一章的父块。

    判据（四条同时成立才提升，避免把正文里提到「第三章」的句子当标题）：
    1. 形态是「第X章 + 短标题」，且标题部分不含句读与虚词；
    2. 文档里已有**被标记**的「第X章」标题 → 本文档确实用这套编号做章节；
    3. 章号与已标记的章号相邻（m±1）→ 是章节序列里缺的那一个；
       同一章号已被标记过 → 是重复项，不提升；
    4. 后面第一个标题是点分编号且首段数字 == 章号 → 「第三章」后面跟着
       「3.1」，章号与节号前缀一致，是强证据（正文里偶然提到「第三章」
       不会满足这条）。
    """
    marked_ords: set[int] = set()
    for i, line in enumerate(lines):
        if i in in_code:
            continue
        m = RE_HEAD.match(line)
        if not m:
            continue
        oc = _chapter_ordinal(m.group(2))
        if oc:
            marked_ords.add(oc[0])
    if not marked_ords:
        return {}

    # 允许的章号：与某个已标记章号相邻（序列里缺的那一个）
    allowed: set[int] = set()
    for m in marked_ords:
        allowed.add(m + 1)
        allowed.add(m - 1)

    found: dict[int, str] = {}
    for i, line in enumerate(lines):
        if i in in_code or RE_HEAD.match(line):
            continue
        s = line.strip()
        if not s or _is_table_row(line):
            continue
        oc = _chapter_ordinal(s)
        if not oc:
            continue
        n, rest = oc
        if n in marked_ords or n not in allowed:
            continue
        if not _title_like(rest):
            continue
        if _next_section_first(lines, i, in_code) != n:
            continue
        found[i] = s
    return found


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


def rebuild_dotted_headings(
    text: str, scheme: LevelScheme | None = None
) -> tuple[str, dict[str, int]]:
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

    角色（而不是绝对级别）
    ----------------------
    本函数判定的是**语义角色**，具体输出几级由 ``scheme`` 决定 ——
    ``scheme`` 来自用户填的父/子分隔符（见 ``heading_levels``）。
    这样插件不必知道用户在分段节点上填了什么分隔符，也不会把级别写死。

    规则
    ----
    * depth1（``1.``）与 depth2（``2.3.``）→ **PARENT**（父块边界）
    * depth3（``2.3.1.``）→ **CHILD**（子块边界）
    * depth4+ → 正文
    * **形态序章级标题**（``第一章 引言`` ``第二章 首页``）→ **PARENT**（见下）
    * 编号章节**内部**的非编号标题（「操作步骤」「合同卡片」「注意:」…）→ 正文
    * 编号章节**之前**的非编号标题（如超大表拆分插入的分组标题）→ 保持原级别

    混合编号体系
    ------------
    用户手册常把「章」写成中文序号、「节」写成点分编号：

        第一章 引言
        1.1. 编写目的
        第二章 首页
        2.1 项目经理工作台

    两套编号混在一篇文档里。``第一章`` 这类**章级**标题必须走独立分支；
    否则它会落进「非编号标题」分支被降成正文 → 父分隔符匹配不到 →
    整章被并进上一个父块。

    混编文档的层级：章 = PARENT，章内点分节 = PARENT，章内子节 = PARENT
    ------------------------------------------------------------------
    混编文档里「章」与「章内的点分节」是**同一粒度的检索单元** ——
    用户要的是「按 3.1、3.2 这样的小节做父分段」。所以章、节、子节
    **同为父块边界**；子分隔符在本分支下无内容可切（父子结构退化为单块，
    检索返回的仍是父块，内容正确）。

    为什么章标题必须与章内节同级：Dify 是**纯文本分隔符切分**，不认标题的
    语义层级。章标题若比节浅一级，它就不含父分隔符 → 永远切不到 → 必然粘在
    **上一个**父块的末尾（实测「# 第二章 首页」落在「1.4 系统业务流程」那个
    父块的末行）。同级后它自己成为父块边界，章标题不再粘尾巴。

    顶层章节（depth1）
    ----------------
    depth1 是顶层章节（``1. 引言`` ``2. 功能说明``），必须自己**开启**一个父块 ——
    降成正文会让它粘在前一节的末尾，检索「1. 引言」时返回的是上一节的内容。

    但顶层章节常常是**空壳**（后面直接跟第一节，没有自己的正文）。空壳标题
    单独占一个父块也不好看 —— 检索命中它只会拿到几个字。所以空壳的 depth1
    与它后面**紧跟的第一个 depth2** 合并成一个父块（这与 Word 分支「父块首行
    是章标题、章的内容跟在后面」的结构一致），depth2 留在同一块内。
    """
    if scheme is None:
        scheme = LevelScheme.default()
    lines = text.split("\n")
    in_code = _code_mask(lines)

    # 上游漏标的「第X章」→ 先恢复成标题行，再走下面的重建逻辑。
    # 漏标的后果很严重：章标题不是标题 → 父分隔符匹配不到 →
    # 整章被并进上一章的父块（实测「第三章 个人中心模块」整章并进了第二章）。
    # 这里先标成 H1 占位，下面的循环会按 RE_SEQ_CHAPTER 统一改写成
    # PARENT 级别，保证「被标记的章标题」与「被恢复的章标题」输出级别完全一致。
    bare = _find_bare_chapters(lines, in_code)
    if bare:
        for i, title in bare.items():
            lines[i] = "# " + title

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
            # 形态序章级标题（第一章/第二章/第三部分）→ 顶层章节。
            #
            # 角色 = PARENT：与章内的点分节**同级**，让「章」和「节」在
            # 同一套分隔符下都成为父块边界（用户要的是「按小节做父分段」）。
            #
            # 为什么章标题要跟小节同级，而不是更高一级：
            # Dify 是**纯文本分隔符切分**，不认标题的语义层级。
            # 若章标题比节浅一级，它就不含父分隔符 → 永远切不到 →
            # 必然粘在**上一个**父块的末尾（实测「# 第二章 首页」落在
            # 「1.4 系统业务流程」那个父块的末行）。
            # 同级后它自己成为父块边界，不再粘尾巴。
            if RE_SEQ_CHAPTER.match(title):
                out.append(scheme.marker(PARENT) + title)
                stats["parent"] += 1
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
            # 空壳顶层章节后的第一个节 → 与它同父块。
            # 角色 = CHILD：它是「章」下面的第一个节，比章更深一级，
            # 但仍留在同一个父块里（不自己开启父块）。
            out.append(scheme.marker(CHILD) + title)
            stats["child"] += 1
            stats["merged"] += 1
            shell_top = False
            continue

        # 混合编号文档：点分编号是「章内」的层级。
        #   第一章 引言          → PARENT（父块边界）
        #   1.1. 编写目的        → PARENT（父块边界，与章标题同级）
        #   1.1.1. 子节           → CHILD （子块边界，与章内节不同级）
        #   第二章 首页          → PARENT（父块边界）
        #   2.1 项目经理工作台    → PARENT（父块边界）
        #
        # 章、节**同为父块边界** —— 父块粒度 = 小节级
        # （用户要的「按 3.1、3.2 做父分段」）。
        #
        # 章内**子节**（depth3）→ CHILD（子块边界），不是 PARENT：
        # 早期把 depth3 也当 PARENT，导致插件把子节也输出成 `## `，
        # 父分隔符一匹配，子节就自己成了一个父块 —— 父块粒度掉到
        # 「子节级」，比用户要的更细（实测「3.1.1 正常立项」独占一个父块）。
        # 降成 CHILD 后子节输出 `### `，正好对上子分隔符：
        # 父块仍是「3.1 交付清单」这一节，子块是节下的各个子节 ——
        # 父子结构才真正建立起来（此前每段只有 1 个子分段，
        # 因为子分隔符一个都匹配不到）。
        #
        # 为什么**不能**降成正文：降成正文会让标题与它的正文落进不同的块 ——
        # 实测「3.1.1 正常立项」成了上一块的末行，正文却成了下一块的首行：
        # 检索命中后模型看到的是标题、没有内容。CHILD 角色保留标记，
        # 标题与正文同在子块内。
        if seq_chapter_seen:
            if depth <= 2:
                out.append(scheme.marker(PARENT) + title)
                stats["parent"] += 1
            elif depth == 3:
                out.append(scheme.marker(CHILD) + title)
                stats["child"] += 1
            else:
                out.append(title)           # depth4+ → 正文
                stats["demoted"] += 1
            continue

        if depth <= 2:
            # depth1（顶层章节）与 depth2（节级）→ PARENT = 父块边界
            out.append(scheme.marker(PARENT) + title)
            stats["parent"] += 1
        elif depth == 3:
            # depth3（小节）→ CHILD = 子块边界
            out.append(scheme.marker(CHILD) + title)
            stats["child"] += 1
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

    ``max_level`` 应传 ``scheme.child_level`` —— 只有与子分隔符对齐的层级
    才需要保留标记，比它更深的一律降回正文（Dify 父子分段只有两级分隔符，
    更深的层级没有对应的切分点，留着只是向量噪音）。
    """
    out = []
    for line in text.split("\n"):
        m = RE_HEAD.match(line)
        if m and len(m.group(1)) > max_level:
            out.append(m.group(2).strip())
        else:
            out.append(line)
    return "\n".join(out)
