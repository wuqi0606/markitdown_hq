# -*- coding: utf-8 -*-
"""md / txt 的标题层级重建（files 分支专用）。

背景
----
插件有两个输入入口，内部链路不同：

* ``files`` —— MarkItDown 解析 → 通用清洗 → 补标题 → （PDF 专项）
* ``text``  —— 通用清洗 → 补标题 → HTML 表格转 MD → **层级重建** → 款级降正文

「层级重建」整段只挂在 ``text`` 入口上（那是 MinerU 解析结果的专用通道），
所以 md / txt 走 ``files`` 入口时会**丢掉这一段**。

后果
----
Dify 父子分段的父分隔符是 ``\\n# ``，**只认一级标题**；检索命中子块后
返回给模型的是**父块全文**。层级没整理 → 父块粒度失控：

* md（MinerU 导出的：封面 H1 + 章节 H2）→ 章节停在 H2 → 父分隔符一个都
  匹配不到 → 整篇挤成 1 个父块；
* txt（纯文本，既无标记也无加粗）→ 标题一个都识别不出来 → 父子结构彻底
  失效（1 个父块 / 1 个子块）。

txt 补不出标题的根因：``heading_restorer`` 的**纯文本路径要求整行加粗**
（``RE_BOLD_TITLE`` 要求 ``**…**``），txt 没有加粗信息 → 一条都补不出来。

判据：编号语义优先
------------------
**不能**用「有没有 H1」当判据 —— MinerU 导出的 md 封面是 H1、章节是 H2，
会被误判成「作者已标注层级」→ 章节停在 H2 → 父分隔符匹配不到 → 照样退化。

真正的判据是**标题文本里有没有编号**：

============================  ==========================================
情形                          处理
============================  ==========================================
无任何标题标记                按编号形态推导层级（形态序 / 点分两套）
标题都无编号                  信任作者标的 ``#`` 级别，只补伪标题
标题有编号                    编号语义优先 → 与 PDF 分支同一套逻辑
============================  ==========================================

「标题有编号」时走 PDF 分支那套，是因为编号本身就是层级语义：
``2.1.1`` 一定比 ``2.1`` 浅，比 ``2.`` 浅 —— 作者标的 ``#`` 级别反而
可能是上游（MinerU）统一标成 ``##`` 的产物，不该采信。

引用关系
--------
* 点分体系 → ``heading_normalizer.normalize_heading_levels``
  + ``pdf_relevel.rebuild_dotted_headings``
* 形态序 → ``pdf_relevel.relevel_headings``
"""

from __future__ import annotations

import re
from collections import Counter

from tools.heading_levels import CHILD, PARENT, LevelScheme
from tools.heading_normalizer import _code_mask, normalize_heading_levels
from tools.heading_restorer import restore_headings
from tools.pdf_relevel import (
    RE_HEAD,
    _is_table_row,
    _seq_form,
    demote_deep_headings,
    rebuild_dotted_headings,
    relevel_headings,
)

# 点分十进制（至少两点）：2.1 / 2.1.1 / 2.1.1.
# 与 pdf_relevel 的判定口径一致，但这里要覆盖**正文行**（txt 还没提升成标题），
# 所以额外放宽「编号后必须带点或空白」的限制 —— 判定体系只需知道「有点分编号」。
RE_DOTTED_ANY = re.compile(r"^(\d+(?:\.\d+)+)\s*(\S.*)$")

# 形态序编号：一、/ （一） / 1、 / （1）
_RE_SEQ_ANY = re.compile(
    r"^\s*(?:"
    r"[一二三四五六七八九十]{1,3}\s*[、.．]"       # 一、
    r"|[（(]\s*[一二三四五六七八九十]{1,3}\s*[)）]"  # （一）
    r"|\d{1,2}\s*[、]"                            # 1、
    r"|[（(]\s*\d{1,2}\s*[)）]"                    # （1）
    r")"
)


def _is_table_row(line: str) -> bool:
    """Markdown 表格行（``| 1 | 20190102 |``）。

    表格里的「1、」「2、」是单元格内容，不是章节编号 —— 判定体系时必须排除。
    """
    s = line.strip()
    return s.startswith("|") and s.endswith("|")


def _text_of(line: str) -> str:
    """取一行的标题文本（去掉 ``#`` 标记）或原文。"""
    m = RE_HEAD.match(line)
    if m:
        return m.group(2).strip()
    return line.strip()


def detect_system(lines: list[str], in_code: set[int] | None = None) -> str | None:
    """判定文档属于哪套编号体系。

    与 PDF 分支的口径一致，但这里要扫**全量行**（含正文）——
    txt 还没提升成标题时，编号只在正文里。

    返回 ``"dotted"`` / ``"seq"`` / ``None``（没有编号体系）。
    """
    has_dotted = False
    has_seq = False
    for i, line in enumerate(lines):
        if in_code and i in in_code:
            continue
        s = _text_of(line)
        if not s or _is_table_row(line):
            continue
        if RE_DOTTED_ANY.match(s):
            has_dotted = True
            break
        if _RE_SEQ_ANY.match(s):
            has_seq = True
    if has_dotted:
        return "dotted"
    if has_seq:
        return "seq"
    return None


def _has_numbered_heading(lines: list[str], in_code: set[int] | None = None) -> bool:
    """是否存在**带编号**的标题行。

    这是「编号语义优先」判据的核心：只看标题文本里有没有编号，
    不看作者标的 ``#`` 级别（上游可能把全部章节统一标成 ``##``）。
    """
    for i, line in enumerate(lines):
        if in_code and i in in_code:
            continue
        m = RE_HEAD.match(line)
        if not m:
            continue
        title = m.group(2).strip()
        if RE_DOTTED_ANY.match(title):
            return True
        if _RE_SEQ_ANY.match(title):
            return True
    return False


def _promote_dotted(
    text: str,
    in_code: set[int] | None = None,
    scheme: LevelScheme | None = None,
) -> tuple[str, dict[str, int]]:
    """无标记的纯文本（点分体系）→ 按编号点数推导标题层级。

    点分体系里「点数 = 层级」是确定的，不需要看任何上下文：

    * depth1 / depth2 → PARENT（父块边界）
    * depth3 → CHILD（子块边界）
    * depth4+ → 正文

    与 ``rebuild_dotted_headings`` 的层级规则保持一致；实际输出几级由
    ``scheme``（来自用户填的父/子分隔符）决定。
    """
    if scheme is None:
        scheme = LevelScheme.default()
    out: list[str] = []
    stats: Counter = Counter()
    for i, line in enumerate(text.split("\n")):
        s = line.strip()
        if not s or _is_table_row(line) or (in_code and i in in_code):
            out.append(line)
            continue
        m = RE_DOTTED_ANY.match(s)
        if not m:
            out.append(line)
            continue
        depth = len(m.group(1).split("."))
        if depth <= 2:
            out.append(scheme.marker(PARENT) + s)
            stats["parent"] += 1
        elif depth == 3:
            out.append(scheme.marker(CHILD) + s)
            stats["child"] += 1
        else:
            out.append(s)               # depth4+ → 正文
            stats["demoted"] += 1
        stats["promoted"] += 1
    stats["system"] = "dotted"
    return "\n".join(out), dict(stats)


class _CodeShield:
    """把代码围栏「整段」藏起来，处理完再放回去。

    ``relevel_headings`` / ``restore_headings`` 都不识别代码围栏 ——
    md 里的 ```python 块常含 ``# 注释`` / ``1. xxx`` 这类行，
    直接放进去会被误当成标题或编号条目。

    实现：把每段代码围栏（``` 标记行 + 内部内容 + ``` 标记行）整体替换成
    **一行**唯一标记，处理完再按出现顺序换回。

    为什么用整段替换而不是逐行占位：``relevel_headings`` 会重排行
    （删空行、加标题标记），行号会整体错位 —— 只能靠**内容顺序**还原。
    整段压成一行后，行号错位就不再影响还原（还原是按出现顺序配对的）。

    标记用纯 ASCII 单词，不匹配 RE_HEAD / RE_DOTTED_ANY / 形态序，
    也不会被 _clean / restore_headings 改写。
    """

    _TOKEN = "codefenceplaceholder"

    def __init__(self, text: str):
        self._lines = text.split("\n")
        self._mask = _code_mask(self._lines)
        self._blocks: list[list[str]] = []      # 每段代码围栏的原始行
        out: list[str] = []
        cur: list[str] = []

        for i, line in enumerate(self._lines):
            if i in self._mask:
                cur.append(line)
                continue
            if cur:
                # 一段围栏结束 → 收进 _blocks，原位压成一行标记
                self._blocks.append(cur)
                out.append(self._TOKEN)
                cur = []
            out.append(line)
        if cur:
            self._blocks.append(cur)
            out.append(self._TOKEN)

        self.text = "\n".join(out)

    def restore(self, text: str) -> str:
        """把标记行换回原来的代码块内容（按出现顺序配对）。"""
        if not self._blocks:
            return text
        lines = text.split("\n")
        result: list[str] = []
        block_i = 0
        for line in lines:
            if line.strip() == self._TOKEN and block_i < len(self._blocks):
                result.extend(self._blocks[block_i])
                block_i += 1
                continue
            result.append(line)
        return "\n".join(result)


def _promote_seq(
    text: str,
    in_code: set[int] | None = None,
    scheme: LevelScheme | None = None,
) -> tuple[str, dict[str, int]]:
    """无标记的纯文本（形态序体系）→ 按编号形态推导标题层级。

    直接复用 PDF 分支的 ``relevel_headings``：它本身就支持从正文提升编号行
    （章 = H1、条 = H2、款 = H3），再把款级降回正文。

    ``relevel_headings`` 不识别代码围栏 —— 代码块先被 ``_CodeShield`` 藏起来，
    处理完再放回去，避免 ``# 注释`` 被当成标题。
    """
    if scheme is None:
        scheme = LevelScheme.default()
    shield = _CodeShield(text) if in_code else None
    work = shield.text if shield else text
    work, stats = relevel_headings(work, scheme)
    work = demote_deep_headings(work, max_level=scheme.child_level)
    if shield:
        work = shield.restore(work)
    stats["system"] = "seq"
    return work, stats


def rebuild_plain_headings(
    text: str, scheme: LevelScheme | None = None
) -> tuple[str, dict[str, int]]:
    """md / txt 的标题层级重建。

    返回 (新文本, 统计)。统计键 ``system`` 取值：

    * ``authored`` —— 标题都无编号，信任作者标的 ``#`` 级别
    * ``dotted``   —— 点分体系，按编号点数重建
    * ``seq``      —— 形态序体系，按编号形态重排
    """
    if scheme is None:
        scheme = LevelScheme.default()
    lines = text.split("\n")
    # 代码围栏内的行不参与标题判定（md 里的 ```python 块常含 `# 注释`）
    in_code = _code_mask(lines)
    heads = [line for i, line in enumerate(lines)
             if i not in in_code and RE_HEAD.match(line)]

    # ---- 情形 1：完全没有标题标记 → 按编号形态推导 ----
    if not heads:
        system = detect_system(lines, in_code)
        if system == "dotted":
            return _promote_dotted(text, in_code, scheme)
        if system == "seq":
            return _promote_seq(text, in_code, scheme)
        # 连编号都没有 → 纯正文，保持原样
        return text, {"system": "plain", "kept": 0}

    # ---- 情形 2：标题都无编号 → 信任作者标的 # 级别 ----
    if not _has_numbered_heading(lines, in_code):
        shield = _CodeShield(text) if in_code else None
        work = shield.text if shield else text
        res = restore_headings(work, scheme=scheme)
        out = shield.restore(res.text) if shield else res.text
        return out, {"system": "authored", "kept": len(heads),
                     "restored": res.added}

    # ---- 情形 3：标题有编号 → 编号语义优先，走 PDF 分支同一套 ----
    system = detect_system(lines, in_code)
    shield = _CodeShield(text) if in_code else None
    work = shield.text if shield else text

    if system == "dotted":
        work, _ = normalize_heading_levels(work)
        work, dstats = rebuild_dotted_headings(work, scheme)
        out = shield.restore(work) if shield else work
        dstats["system"] = "dotted"
        return out, dstats

    work, rstats = relevel_headings(work, scheme)
    work = demote_deep_headings(work, max_level=scheme.child_level)
    out = shield.restore(work) if shield else work
    rstats["system"] = "seq"
    return out, rstats
