# -*- coding: utf-8 -*-
"""还原 Word 自动编号列表项的真实编号。

问题
----
MarkItDown（mammoth）转 docx 时，**把不同编号体系的列表合并重排**成一条
连续的有序列表。源文档里明明是四套独立编号：

    numId=3  lvlText='%1.'   →  1.  2.  3.  4.
    numId=4  lvlText='（%1）' →  （1）（2）
    numId=5  lvlText='(%2)'  →  (1) (2)
    numId=6  lvlText='(%2)'  →  (1) (2) (3) (4)

mammoth 全部渲染成 ``1. 2. 3. 4. ...``，于是序号和形态全错：

    2. 提前实施项目及试用项目注销。…   ← 原文是「（1）」
    3. 正常实施项目注销。…              ← 原文是「（2）」

正文里写「详见（1）」时，检索片段对不上号 —— 知识库里的编号是假的。

解法
----
编号的真相在 OOXML 里：``numId`` / ``ilvl`` 指向 ``numbering.xml`` 的
``lvlText`` 模板（如 ``（%1）``），配合每个段落自己的计数快照即可还原出
「这一项在原文里到底长什么样」。

只处理**自动编号**的段落（``numId`` 存在）。手打编号（段落文本里本来就写着
「1. xxx」）的编号在文本里，markdown 里已经是对的，不动。

输出形态
--------
**始终保持 markdown 有序列表语法**（``1. 正文``），只把序号换成 XML 里的真值。

为什么不还原成原文的「（1）」「1、」形态？
    按 CommonMark 解析，``(1)正文`` 不是列表项，只是普通文本 ——
    缩进续行还会被合并成**一个 paragraph**，整组列表塌成一段文字，
    直接毁掉分段结构。``1. 正文`` 才是独立的 list_item。

序号本身必须还原：
    mammoth 把四套独立编号合并重排成连续序号后，
    正文里「详见（1）」这种交叉引用就再也对不上号了。
    序号错了比形态错了危害大得多。
"""

from __future__ import annotations

import re

from tools.docx_meta import DocxMeta, ParaMeta, render_lvl_text

# markdown 列表项：可选 bullet 前缀 + 数字 + 分隔符 + 正文
RE_ITEM = re.compile(
    r"^(?P<ind>[ \t]*)(?P<bullet>[*+-][ \t]+)?"
    r"(?P<num>\d{1,3})(?P<sep>[.)])[ \t]+(?P<rest>.*)$"
)


def render_prefix(p: ParaMeta, meta: DocxMeta) -> str:
    """自动编号段落还原出来的原始编号前缀，如「一、」「（1）」「1.」。

    用段落**当时**的计数快照（``p.counters``），不是最终计数 ——
    否则一组标题会全渲染成同一个编号。
    """
    if not (p.is_auto_numbered and p.lvl_text and p.num_id):
        return ""
    counters = p.counters or meta.num_counters.get(p.num_id, {})
    return render_lvl_text(p.lvl_text, counters, p.num_fmt, p.ordinal or 1)


def true_ordinal(p: ParaMeta, meta: DocxMeta) -> int | None:
    """这一项在源文档里的**真实序号**（阿拉伯数字）。

    从编号模板渲染出的前缀里提数字：``（1）`` → 1、``1、`` → 1、
    ``第三章`` → 3。提不出数字（纯中文数字编号、字母编号）时返回 None ——
    那类序号在 markdown 里本来也不适用，不动。
    """
    prefix = render_prefix(p, meta)
    if not prefix:
        return None
    nums = re.findall(r"\d+", prefix)
    if not nums:
        return None
    return int(nums[-1])


def restore_item_prefix(line: str, p: ParaMeta | None, meta: DocxMeta) -> str:
    """把一个 markdown 列表项的序号换成源文档里的真实序号。

    只换**序号**（``3.`` → ``1.``），保持 markdown 有序列表语法 ——
    还原成「（1）」「1、」这类形态会让列表项塌成普通段落。

    `p` 为 None（对齐不上）或段落不是自动编号时原样返回。
    """
    if p is None or not p.is_auto_numbered:
        return line
    if p.num_fmt == "bullet":
        # 本来就是无序列表，mammoth 的 `*` 是对的
        return line
    m = RE_ITEM.match(line)
    if m is None:
        return line
    num = true_ordinal(p, meta)
    if num is None:
        return line
    if num == int(m.group("num")):
        return line

    ind = m.group("ind")
    bullet = m.group("bullet") or ""
    rest = m.group("rest")
    return f"{ind}{bullet}{num}. {rest}"
