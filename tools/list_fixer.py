# -*- coding: utf-8 -*-
"""修复 mammoth 列表降级造成的畸形嵌套。

现象
----
Word 的自动编号列表里，如果某一层是「孤立」的（该 numId 下没有更浅的层级
可挂靠），mammoth 会给它补一个虚拟父级 ``<ul>``，把真正的有序列表套进去：

    <ul><li><ol><li>费用已结清的…</li><li>费用未结清的…</li></ol></li></ul>

markdownify 忠实翻译，于是 Markdown 里多出一个 ``*``：

    * 1. 费用已结清的，即项目所有物资完成返库，…
      2. 费用未结清的，即合同不合作，…

两个害处：
1. 列表语义被破坏 —— 明明是有序编号，读起来像无序列表；
2. 按 CommonMark 解析，续行缩进不足 3 空格，会被当成上一项的**惰性续行**，
   整组列表塌成一段文字。

判定
----
只认一种形态：head 行是 ``* N. …``（bullet + 数字序号），紧跟至少一行
缩进的 ``N. …``（**行首没有列表符号**）。

真正的 bullet 列表每行都带 ``*``，续行同样带 ``*`` —— 不会被命中。
再要求续行编号从 ``head_num + 1`` 起连续递增（markdownify 的 ``<ol>``
一定从 1 开始连续），避免误伤「文本本身以 `1. ` 开头 + 嵌套列表」这种
合法结构。

XML 交叉验证（meta 可用时）
--------------------------
docx 元数据给出该段落的 ``numId`` 与 ``numFmt``。若 md 行对应的是
**自动编号的有序**列表项（``numId`` 存在且 ``numFmt`` 不是 bullet），
说明它本该是 ``<ol>``、被 mammoth 降级成了 ``<ul>`` —— 确认是缺陷，
必须修。反之若编号是手打文本（``numId`` 为空）或本来就是 bullet，就不动。
"""

from __future__ import annotations

import re

from tools.docx_meta import DocxMeta
from tools.md_aligner import align

# head 行：`* 1. 内容` / `+ 1. 内容` / `- 1. 内容`
_RE_HEAD = re.compile(
    r"^(?P<ind>[ \t]*)[*+-][ \t]+(?P<num>\d{1,3})(?P<sep>[.)])[ \t]+(?P<rest>\S.*)$"
)
# cont 行：缩进的 `2. 内容`，行首没有列表符号
_RE_CONT = re.compile(
    r"^(?P<ind>[ \t]+)(?P<num>\d{1,3})(?P<sep>[.)])[ \t]+(?P<rest>\S.*)$"
)


def _cont_is_sequential(head_num: int, conts: list[re.Match[str]]) -> bool:
    """续行编号必须从 head_num + 1 起连续递增。"""
    expect = head_num + 1
    for cm in conts:
        if int(cm.group("num")) != expect:
            return False
        expect += 1
    return True


def _xml_confirms_defect(line_no: int, alg) -> bool | None:
    """XML 是否确认这是「有序列表被降级成 bullet」。

    返回 True/False；None 表示对齐不上、无法判断。
    """
    p = alg.get(line_no)
    if p is None:
        return None
    if p.num_id is None:
        # 编号是手打文本（段落里本来就写着「1. 」），不是自动编号 → 不动
        return False
    if p.num_fmt == "bullet":
        # 本来就是无序列表 → 不动
        return False
    return True


def fix_lists(text: str, meta: DocxMeta | None = None) -> tuple[str, int]:
    """修复畸形嵌套列表，返回 (新文本, 修复处数)。

    `meta` 为 docx 元数据；传入时用 XML 交叉验证判定，None 时只靠文本形态。
    """
    if not text:
        return text, 0

    lines = text.split("\n")
    alg = align(text, meta) if meta is not None else None

    out: list[str] = []
    fixed = 0
    i = 0
    n = len(lines)
    while i < n:
        m = _RE_HEAD.match(lines[i])
        if m is None:
            out.append(lines[i])
            i += 1
            continue

        # 收集紧跟其后的缩进数字行（无列表符号）
        conts: list[re.Match[str]] = []
        j = i + 1
        while j < n:
            cm = _RE_CONT.match(lines[j])
            if cm is None:
                break
            conts.append(cm)
            j += 1

        if not conts or not _cont_is_sequential(int(m.group("num")), conts):
            out.append(lines[i])
            i += 1
            continue

        # XML 交叉验证：确认是「有序被降级」而非「文本自带编号」
        if alg is not None:
            confirmed = _xml_confirms_defect(i, alg)
            if confirmed is False:
                out.append(lines[i])
                i += 1
                continue

        # ---- 修复：去掉虚拟父级的 bullet 标记，续行缩进同步回收 ----
        base = m.group("ind")
        sep = m.group("sep")
        out.append(f"{base}{m.group('num')}{sep} {m.group('rest')}")

        k = i + 1
        for cm in conts:
            ind = cm.group("ind")
            extra = ind[len(base):] if ind.startswith(base) else ind.lstrip()
            # markdownify 把嵌套 <ol> 缩进 2 格（= bullet 标记宽度），回收掉
            if extra.startswith("  "):
                extra = extra[2:]
            elif extra.startswith("\t"):
                extra = extra[1:]
            out.append(f"{base}{extra}{cm.group('num')}{sep} {cm.group('rest')}")
            k += 1

        fixed += 1
        i = k

    return "\n".join(out), fixed
