# -*- coding: utf-8 -*-
"""补齐 Markdown 中丢失的标题层级。

背景：MarkItDown 靠 Word 的「标题1/2/3」样式映射生成 `#`。
企业制度文件大量章节标题用的是**普通段落**（手打编号，或 Word 自动编号），
转换后层级完全丢失，父子分段的 `\\n# ` / `\\n## ` 分隔符随之失效。

标题分三类，信号强度递减：

**R1 真实标题** —— mammoth 已经渲染成 `#` 的段落。
    层级直接用 md 里的 `#` 数量（与 Word 样式表的 outlineLvl 一致）。
    额外做两件事：去掉标题里多余的 `**`；把它的**编号形态**登记进作用域。

**R2 强真标题** —— 作者设了大纲级别（`w:outlineLvl`）但 mammoth 没认出来的段落。
    实测这类段落有 278 个被完全漏掉：样式名不是 `heading N`（中文 Word 常见），
    或者作者直接在段落上设了大纲级别。
    层级 = outlineLvl + 1（作者权威）；但要用字体/字号复核 ——
    作者常把正文段落也设成大纲级别（「1、负责仓库管理制度的编制和优化」）。

**R3 伪标题** —— 靠字号/字体/加粗/编号组投票推断出来的标题。
    层级由**作用域栈**推断，不靠全局猜形态序。

层级模型用**作用域栈**，维护一条「当前打开的章节链」：

    map: 编号形态 → 层级
    遇到标题形态 F：
      - 关闭比 F 更细的形态（那些章节已经结束）
      - 若 F 已在链上 → 同一层级（兄弟）
      - 否则 → 比链上最深的形态深一层
    遇到权威标题（R1/R2）：
      - 登记形态 → 层级（层级由 XML 决定，不由形态决定）

这样「（一）项目注销」(H2) 下面的「1. 发起申请」会自动变成 H3，
而同一章里的第二个「1. xxx」还是 H3 —— 不需要全局猜形态序。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from tools.docx_meta import (
    FORM_RANK,
    DocxMeta,
    ParaMeta,
    canonical_form,
    form_from_text,
    render_lvl_text,
)
from tools.heading_levels import CHILD, DEEP, PARENT, LevelScheme
from tools.num_prefix import restore_item_prefix

# ---- 标题识别规则（纯文本路径用） ----
RE_CHAPTER = re.compile(r"^\s*第[一二三四五六七八九十百]+\s*[章节]\s*\S")
RE_ARTICLE = re.compile(r"^\s*第[一二三四五六七八九十百]+\s*条\s*\S")
RE_CN = re.compile(r"^\s*[一二三四五六七八九十]+\s*[、.．]\s*\S")
RE_NUM = re.compile(r"^\s*\d{1,2}\s*[、.．]\s*\S")
RE_ATTACH = re.compile(r"^\s*附\s*件\s*[一二三四五六七八九十\d]+\s*[、.．]?\s*\S")
RE_CN_PAREN = re.compile(r"^\s*[（(]\s*[一二三四五六七八九十]+\s*[)）]\s*\S")
RE_NUM_PAREN = re.compile(r"^\s*[（(]\s*\d{1,2}\s*[)）]\s*\S")

FORMS = [
    ("chapter", RE_CHAPTER),
    ("cn", RE_CN),
    ("article", RE_ARTICLE),
    ("cnp", RE_CN_PAREN),
    ("num", RE_NUM),
    ("nump", RE_NUM_PAREN),
]
_ATTACH_ALIAS = "attach"

RE_EXISTING_H = re.compile(r"^\s*(#{1,6})\s")
RE_EMPTY_HEADING = re.compile(r"^\s*#{1,6}\s*$")
RE_FM_DELIM = re.compile(r"^\s*---\s*$")
RE_TOC_MARK = re.compile(r"^\s*(?:#{1,6}\s*)?目\s*录\s*[:：]?\s*$")
RE_TOC_TAIL = re.compile(r".*\s\d{1,3}$")
RE_ANNOTATION = re.compile(r"^\s*[（(][^)）]{0,60}[)）]\s*$")

RE_BOLD_TITLE = re.compile(
    r"^\s*(?P<pre>(?:\d{1,2}\s*[、.．]\s*)?)\*\*(?P<title>.+?)\*\*(?P<post>.*)$"
)

MAX_TITLE_LEN = 45
BAD_CHARS = set("。；！？，")
SENT_END = set("。；！？，、；")
# 含句读标点就不是标题（「第七条 …（1、…；）」以右括号收尾，末尾检查抓不到）
SENT_PUNCT = set("。；！？，")


@dataclass
class HeadInfo:
    """一个标题的判定结果（供回归脚本核验用）。"""

    level: int
    body: str
    kind: str                # real | outlined | pseudo
    form: str | None = None
    line: int = 0            # 在输出文本中的行号
    parent: int | None = None  # 父章节实例号（判断兄弟关系用）


@dataclass
class RestoreResult:
    """补齐结果。"""

    text: str
    added: int = 0
    heads: list[HeadInfo] = field(default_factory=list)
    renumbered: int = 0        # 还原掉 mammoth 错误序号的列表项数


def strip_bold(s: str) -> str:
    """去掉 markdown 加粗/斜体标记。"""
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"__(.+?)__", r"\1", s)
    s = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"\1", s)
    return s.strip()


def _bold_parts(raw: str) -> tuple[str, str, str] | None:
    """拆解「加粗标题行」→ (编号前缀, 加粗标题文本, 加粗之后的补充说明)。"""
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
    """返回标题文本命中的编号形态 key；不是候选标题返回 None。"""
    if not text or len(text) > MAX_TITLE_LEN:
        return None
    if BAD_CHARS & set(text):
        return None
    if RE_TOC_TAIL.match(text):
        return None
    if RE_ATTACH.match(text):
        return _ATTACH_ALIAS
    for key, pat in FORMS:
        if pat.match(text):
            return key
    return None


def _is_annotation(post: str) -> bool:
    """标题后面跟的内容，是不是「补充说明」而不是正文。

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
    """判定该行是不是「加粗的编号标题」——真伪标题（纯文本路径）。"""
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
    也会让 `\\n# ` 分隔符在带粗体的标题上匹配不到。
    """
    m = RE_EXISTING_H.match(raw)
    if not m:
        return raw
    hashes = m.group(1)
    text = raw[m.end():]
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"__(.+?)__", r"\1", text)
    return f"{hashes} {text.strip()}"


# ------------------------------------------------------------------ 作用域栈


class Scope:
    """当前打开的章节链：编号形态 → 层级。

    形态序（从粗到细）：
        chapter(第x章/第x部分) < cn(一、) < article(第x条) < cnp((一)) < num(1、) < nump((1))

    链上「形态越粗 → 层级越浅」。进入某个形态时，**比它更细的形态全部关闭**
    （那些章节已经结束）；比它更粗的父章节保留。
    这样「一、xxx」下面的「（一）xxx」能自然落到下一层，
    而遇到下一个「二、xxx」时又能正确地回到同一层。
    """

    def __init__(self) -> None:
        self.map: dict[str, int] = {}   # 打开的章节 form -> level
        self.ident: dict[str, int] = {}  # form -> 章节实例号（每次重开就换号）
        self._seq = 0
        self.real_floor: int = 0         # 最近一个权威标题的层级

    def _renew(self, form: str) -> None:
        """给形态换一个实例号 —— 表示「这一段章节是新的一个」。"""
        self._seq += 1
        self.ident[form] = self._seq

    def _pop_finer(self, form: str) -> None:
        """关闭比该形态更细的章节（那些章节已经结束）。"""
        rank = FORM_RANK.get(form, 0)
        for f in list(self.map):
            if FORM_RANK.get(f, 0) > rank:
                del self.map[f]
                self.ident.pop(f, None)

    def _coarser_max(self, form: str) -> int | None:
        """链上比该形态更粗的章节里，最深的那个层级。"""
        rank = FORM_RANK.get(form, 0)
        lv = [lvl for f, lvl in self.map.items()
              if FORM_RANK.get(f, 0) < rank]
        return max(lv) if lv else None

    def parent_ident(self, form: str | None) -> int | None:
        """父章节的实例号 —— 用来判断两个同形态标题是不是兄弟。"""
        if not form:
            return None
        rank = FORM_RANK.get(form, 0)
        best = None
        for f in self.map:
            if FORM_RANK.get(f, 0) < rank:
                if best is None or FORM_RANK.get(f, 0) > FORM_RANK.get(best, 0):
                    best = f
        return self.ident.get(best) if best else None

    def level_for(self, form: str) -> int:
        """推断标题的层级。"""
        self._pop_finer(form)
        if form in self.map:
            # 该形态的章节还开着 → 同一层级（兄弟）
            lvl = self.map[form]
        else:
            # 新开一层：比链上更粗的最深章节深一层
            coarser = self._coarser_max(form)
            lvl = (coarser + 1) if coarser is not None else self.real_floor + 1
            self._renew(form)
        lvl = min(lvl, 6)
        self.map[form] = lvl
        return lvl

    def bind(self, form: str | None, level: int) -> None:
        """登记一个真实标题（md 里的 `#`，层级由 Word 标题样式决定）。"""
        self.real_floor = level
        if form:
            self._pop_finer(form)
            self.map[form] = level
            self._renew(form)
        else:
            # 形态未知：关闭同层及更深的章节，父章节保留
            for f in list(self.map):
                if self.map[f] >= level:
                    del self.map[f]
                    self.ident.pop(f, None)

    def note(self, level: int) -> None:
        """登记一个「强真标题」的层级（XML 大纲级别），作为上下文深度基准。"""
        self.real_floor = level


# ------------------------------------------------------------------ XML 判定


def _short_title(text: str, limit: int = 24) -> bool:
    """短标题形态：不太长、不以句读结尾。"""
    t = text.strip()
    if not t or len(t) > limit:
        return False
    return not (SENT_END & set(t[-1:]))


def _is_head_font(p: ParaMeta, meta: DocxMeta) -> bool:
    """段落的字体是不是标题字体。"""
    return bool(p.font) and p.font in meta.head_fonts


def _is_bigger(p: ParaMeta, meta: DocxMeta) -> bool:
    """段落的字号是不是比正文大。"""
    if p.size is None or meta.base_size is None:
        return False
    return p.size > meta.base_size


def _para_is_outlined(p: ParaMeta, meta: DocxMeta) -> bool:
    """是不是「作者设了大纲级别、但 mammoth 没转成 `#`」的强真标题。

    强真标题还要过一道复核：作者常把正文段落也设成大纲级别
    （「1、负责仓库管理制度的编制和优化」就是 out=2 的正文）。
    复核靠**字体/字号**——标题用黑体/楷体，正文用仿宋/宋体。
    另外要求是「短标题」形态：长句即便标了大纲级别也当正文。
    """
    if p.outline is None:
        return False
    if p.in_table or p.in_toc or p.in_cover:
        return False
    text = p.text.strip()
    if not text:
        return False
    if not _short_title(text):
        return False
    if SENT_PUNCT & set(text):
        return False
    # 字体 / 字号复核
    if _is_head_font(p, meta) or _is_bigger(p, meta):
        return True
    return False


def _para_is_pseudo(p: ParaMeta, meta: DocxMeta) -> bool:
    """是不是伪标题（字号/字体/加粗/编号组投票推断出来的）。"""
    if p.in_table or p.in_toc or p.in_cover:
        return False
    if p.outline is not None:
        # 有大纲级别的走强真标题通道
        return False

    if p.is_auto_numbered:
        # 自动编号：整组投票（组内全是短标题 → 整组是标题）
        if p.key in meta.heading_groups:
            return True
        # 单条组要求加粗
        return bool(p.bold) and _short_title(p.text)

    # 手打编号的段落
    if p.form is None:
        return False
    text = p.text.strip()
    if not text or len(text) > MAX_TITLE_LEN:
        return False
    if SENT_PUNCT & set(text):
        return False
    if _is_bigger(p, meta):
        return True
    if _is_head_font(p, meta):
        return True
    if p.bold:
        return True
    return False


def _para_form(p: ParaMeta, meta: DocxMeta) -> str | None:
    """docx 段落的编号形态（用于层级映射）。"""
    if p.is_auto_numbered:
        # 优先用段落自身的形态；没有就用所属编号组的形态
        if p.form is not None:
            return p.form
        return meta.group_form.get(p.key)
    return canonical_form(p.form)


def _number_prefix(p: ParaMeta, meta: DocxMeta) -> str:
    """自动编号段落还原出来的原始编号前缀，如「一、」「（三）」。"""
    if not (p.is_auto_numbered and p.lvl_text and p.num_id):
        return ""
    counters = p.counters or meta.num_counters.get(p.num_id, {})
    return render_lvl_text(p.lvl_text, counters, p.num_fmt, p.ordinal or 1)


def _heading_text(p: ParaMeta, meta: DocxMeta) -> str:
    """标题正文。自动编号的段落把原始编号还原回来。"""
    body = p.text.strip()
    prefix = _number_prefix(p, meta)
    if prefix:
        body = f"{prefix}{body}"
    return body


def _restore_item_prefix(line: str, p: ParaMeta | None,
                         meta: DocxMeta) -> tuple[str, bool]:
    """还原普通列表项的编号（不是标题的那些）。

    mammoth 把不同编号体系的列表**合并重排**成连续序号，
    「（1）」「1、」「(1)」全被渲染成 `1.` `2.` `3.`。
    这里用 XML 的编号模板还原回去。
    """
    if p is None:
        return line, False
    new = restore_item_prefix(line, p, meta)
    return new, new != line


def _md_text_of(line: str) -> str:
    """md 行的纯文本（去掉标题标记与加粗）。"""
    m = RE_EXISTING_H.match(line)
    if m:
        return strip_bold(line[m.end():]).strip()
    return strip_bold(line).strip()


# ------------------------------------------------------------------ 主流程


def _level_marker(lvl: int, scheme: LevelScheme) -> str:
    """Word 标题层级 → Markdown 标记。

    Word 的标题样式层级（Heading1/2/3…）先映射成语义角色，再由 ``scheme``
    换算成实际输出的 ``#`` 数量：

    * 第 1 层（Heading1，章）→ PARENT → 父分隔符对应的级别
    * 第 2 层（Heading2，条）→ CHILD  → 子分隔符对应的级别
    * 第 3 层及更深（Heading3，款）→ DEEP → 比子块更深一级

    默认 scheme（父 1 级 / 子 2 级）下与历史行为逐字节一致。
    """
    if lvl <= 1:
        return scheme.marker(PARENT)
    if lvl == 2:
        return scheme.marker(CHILD)
    return scheme.marker(DEEP, lvl - 2)


def restore_headings(
    text: str,
    meta: DocxMeta | None = None,
    scheme: LevelScheme | None = None,
) -> RestoreResult:
    """补齐标题层级。

    `meta` 为 docx 元数据（由 `docx_meta.parse_docx_meta` 产出）。
    传入时用 XML 信号判定伪标题；为 None 时退回到纯文本规则。

    `scheme` 决定输出的**绝对级别**：Word 的标题样式层级（Heading1/2/3）
    映射成语义角色（第 1 层 = PARENT、第 2 层 = CHILD、更深 = DEEP），
    再由 scheme 换算成实际输出的 ``#`` 数量 —— 这样插件不必知道用户在
    分段节点上填了什么分隔符（见 `heading_levels`）。
    """
    if scheme is None:
        scheme = LevelScheme.default()
    lines = text.split("\n")

    scope = Scope()
    out: list[str] = []
    heads: list[HeadInfo] = []
    in_fm = False
    in_toc = False
    added = 0
    renumbered = 0
    seen: set[str] = set()

    if meta is not None and meta.paras:
        from tools.md_aligner import align
        alg = align("\n".join(lines), meta)
    else:
        alg = None

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

        # ---- R1 真实标题（markitdown 生成的） ----
        existing = _real_heading_level(line)
        if existing is not None:
            p = alg.get(i) if alg else None
            form = _para_form(p, meta) if (p is not None and meta is not None) else None
            if form is None:
                form = canonical_form(_form_of(_md_text_of(line)))
            scope.bind(form, existing)
            body = _normalize_heading(line)
            # Word 的自动编号由 Word 渲染，mammoth 不写进文本 ——
            # 从 XML 还原回来，标题才和目录一致（「编制目的」→「一、编制目的」）
            if p is not None and meta is not None:
                prefix = _number_prefix(p, meta)
                if prefix and not _md_text_of(body).startswith(prefix):
                    body = f"{_level_marker(existing, scheme)}{prefix}{_md_text_of(body)}"
            else:
                # 已有标记的标题：按语义层级重新编号（scheme 决定实际级别）
                body = f"{_level_marker(existing, scheme)}{_md_text_of(body)}"
            out.append(body)
            heads.append(HeadInfo(existing, _md_text_of(body), "real", form,
                                  len(out) - 1, scope.parent_ident(form)))
            continue

        # ---- XML 元数据路径 ----
        if alg is not None and meta is not None:
            p = alg.get(i)
            if p is not None:
                form = _para_form(p, meta)
                # R2 强真标题：作者设了大纲级别，mammoth 没转
                if _para_is_outlined(p, meta):
                    # 层级同样走章节链 —— 作者的 outlineLvl 偶尔自相矛盾
                    # （同一章里「一、xxx」标 outline=0、「二、xxx」标 outline=1），
                    # 章节链按形态序推断更符合「兄弟同级」。
                    # 注意：不抬 real_floor —— 层级已由章节链决定，
                    # 再抬一次会自我放大（第X部分会一路涨到 H5）。
                    lvl = (scope.level_for(form) if form
                           else min(p.outline + 1, 6))
                    body = _heading_text(p, meta)
                    if body and body not in seen:
                        seen.add(body)
                        out.append(_level_marker(lvl, scheme) + body)
                        added += 1
                        heads.append(HeadInfo(lvl, body, "outlined", form,
                                              len(out) - 1, scope.parent_ident(form)))
                        continue
                # R3 伪标题
                if _para_is_pseudo(p, meta):
                    if form is not None:
                        lvl = scope.level_for(form)
                        body = _heading_text(p, meta)
                        if body and body not in seen:
                            seen.add(body)
                            out.append(_level_marker(lvl, scheme) + body)
                            added += 1
                            heads.append(HeadInfo(lvl, body, "pseudo", form,
                                          len(out) - 1, scope.parent_ident(form)))
                            continue
            # 不是标题的普通行：还原 mammoth 合并重排掉的编号
            # （原文的「（1）」「1、」「(1)」被渲染成了 `1.` `2.` `3.`）
            line, fixed = _restore_item_prefix(line, p, meta)
            if fixed:
                renumbered += 1
            out.append(line)
            continue

        # ---- 纯文本路径 ----
        cls = _classify(line)
        if cls is None:
            out.append(line)
            continue

        key, body, post = cls
        lvl = scope.level_for(key)
        if body not in seen:
            seen.add(body)
            out.append(_level_marker(lvl, scheme) + body)
            added += 1
            heads.append(HeadInfo(lvl, body, "pseudo", key,
                                          len(out) - 1, scope.parent_ident(key)))
            if post:
                out.append("")
                out.append(post)
        else:
            out.append(line)

    result = re.sub(r"\n{3,}", "\n\n", "\n".join(out))
    return RestoreResult(text=result, added=added, heads=heads,
                         renumbered=renumbered)
