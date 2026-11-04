# -*- coding: utf-8 -*-
"""剥掉标题行里的章节编号。

为什么
----
标题里的编号（「2.7.1.」「一、」「第一章」）在向量检索里是纯噪声：
它们既不携带语义，又把标题词与查询词之间的语义距离拉开。
实测同一段内容，搜「合同签订」（父块标题词）得分 0.7267，
搜「合同起草」（子块标题词）只有 0.5791 —— 差的 0.15 主要就来自
「2.7.1.」这几个字符占掉了标题的注意力。

只剥标题行
----------
正文里的编号（「1-指定条件」「①签订日期：」「(1) …」）是**内容的一部分**，
不能动 —— 只有以 ``#`` 开头的标题行才处理。

什么时候剥
----------
必须在**标题层级重建之后**剥。层级判定依赖编号形态
（形态序 / 点分十进制体系靠编号点数与形态识别深度），
先剥编号会让层级判定失去依据。

副作用
------
按编号定位的能力会失效：搜「2.7.1」不再命中。这是显式的取舍，
换来的是「合同起草」这类**语义词**能召回。
"""

from __future__ import annotations

import re

from tools.heading_normalizer import _code_mask

# 标题行：# 后必须有空白，避免把 `#标签` 当标题
_RE_HEADING = re.compile(r"^(#{1,6})[ \t]+(.*)$")

# ---------------------------------------------------------------- 编号形态
#
# 判定原则：**编号与标题文字之间必须有分隔**（点+空白 / 纯空白 / 括号）。
# 否则「3.5万元」这类标题会被误剥 —— 那是金额，不是章节号。
# 这一点与插件自身的标题判定口径一致（点分体系的深度识别同样要求分隔处）。

# 1) 点分十进制：2.7.1. / 2.7.1 / 2.7.1（编号后直接接文字）
#    「3.5万元」不匹配 —— 编号后没有分隔符。
_RE_DOTTED_MULTI = re.compile(r"^\d+(?:\.\d+)+(?:\.[ \t　]*|[ \t　]+)")

# 1b) 单层编号 + 点：1. 合同起草 / 1.合同起草
#     点后可以没有空白（`1.合同起草` 是常见写法），
#     但点后**不能紧跟数字** —— 否则「3.5 万元」会被剥成「5 万元」。
_RE_DOTTED_ONE = re.compile(r"^\d+\.(?!\d)[ \t　]*")

# 3) 中文章节：第一章 / 第二节 / 第十三条
_RE_CN_CHAPTER = re.compile(r"^第[一二三四五六七八九十百零〇\d]+[章节条部篇][ \t　]*")

# 4) 形态序条目：一、 / 二、
_RE_CN_SEQ = re.compile(r"^[一二三四五六七八九十]{1,3}[、.．][ \t　]*")

# 5) 阿拉伯数字条目：1、 / 2.
#    数字后面不能再跟数字 —— 否则「3.5万元」会被当成「3.」剥掉。
_RE_NUM_SEQ = re.compile(r"^\d{1,2}[、.．](?!\d)[ \t　]*")

# 6) 括号条目：（一） / (1)
_RE_PAREN = re.compile(r"^[（(][一二三四五六七八九十\d]{1,3}[)）][ \t　]*")

# 7) 附件编号：附件一、 / 附件 1
#    「附件」标题**不剥** —— 附件号是定位标识（正文会写「见附件一」），
#    剥掉后标题变成「附件 ：」，既丢了标识又留下多余空格。

# 按顺序尝试，剥掉第一个匹配到的编号前缀。
# 点分十进制优先 —— 「2.7.1.」比「2.」更具体，先剥完整的。
_PATTERNS = (
    _RE_DOTTED_MULTI,
    _RE_DOTTED_ONE,
    _RE_CN_CHAPTER,
    _RE_PAREN,
    _RE_CN_SEQ,
    _RE_NUM_SEQ,
)

def strip_heading_numbers(text: str) -> tuple[str, int]:
    """剥掉标题行开头的章节编号。

    返回 (新文本, 剥掉的标题数)。非标题行一律不动。
    """
    if not text:
        return text, 0

    lines = text.split("\n")
    in_code = _code_mask(lines)
    out = list(lines)
    changed = 0

    for i, line in enumerate(lines):
        if i in in_code:
            continue
        m = _RE_HEADING.match(line)
        if not m:
            continue
        marks, title = m.group(1), m.group(2)
        if not title.strip():
            continue
        for pat in _PATTERNS:
            new = pat.sub("", title, count=1)
            if new != title:
                # 剥完必须还剩下有效标题文字，否则说明编号本身就是标题
                # （如「## 3.5」这种纯数字标题）—— 保留原样。
                if new.strip():
                    out[i] = f"{marks} {new.strip()}"
                    changed += 1
                break

    return "\n".join(out), changed
