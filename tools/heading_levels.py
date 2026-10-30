# -*- coding: utf-8 -*-
"""标题级别 ← 分段器分隔符 的双向映射。

背景
----
Dify 的父子分段（``parentchild_chunker`` 插件）用两个**纯文本分隔符**切块：

* ``separator``（父块分隔符）→ 把整篇切成父块；
* ``subchunk_separator``（子块分隔符）→ 在每个父块内切成子块。

切分是 ``text.split(分隔符)``，**不认标题的语义层级** —— 只有行首那串
``#`` 恰好等于分隔符里的那串，标题才成为块边界。

于是「插件输出几级标题」和「Dify 侧填什么分隔符」必须是**同一个约定**，
否则会出现两种典型故障：

* 插件输出 ``# X``，分隔符填 ``\\n## `` → ``\\n## `` 切不到 ``# X``
  → 该标题粘在**上一个**父块的末尾，章节被腰斩；
* 插件输出 ``## X``，分隔符填 ``\\n# `` → 父分隔符先命中
  （``\\n# `` 是 ``\\n## `` 的前缀，两级分隔符互相打架）→ 父子结构错乱。

所以插件不能再把级别写死。它需要知道用户在分段节点上填了什么。

分隔符 → 级别
--------------
分隔符里的 ``#`` 个数就是它对应的标题级别：

=================  ======  ==================================
分隔符              级别    插件应当输出的形式
=================  ======  ==================================
``\\n# ``            1      ``# 标题``
``\\n## ``           2      ``## 标题``
``\\n### ``          3      ``### 标题``
=================  ======  ==================================

分隔符里**没有** ``#``（例如 Dify 的默认 ``\\n\\n`` / ``\\n``）时，
切分与标题标记无关 —— 此时退回默认约定（父 1 级 / 子 2 级）。

解码
----
Dify 侧对分隔符做了 ``codecs.decode(sep, "unicode_escape")``
（见 ``parent_child_chunk`` 插件的 ``FixedRecursiveCharacterTextSplitter``），
所以用户填的 ``\\n# `` 在切分时是「换行 + ``# ``」。本模块用同样的方式
解码，保证两边对分隔符的理解一致。

等级角色
--------
插件内部各分支判定出的是**角色**（而不是绝对级别）：

============  ================================================
角色           含义
============  ================================================
``PARENT``    顶层章节 —— 应当成为父块边界
``CHILD``     节 —— 应当成为子块边界
``DEEP``      更深的层级 —— 保留标记，但不与任何分隔符对齐
``BODY``      正文 —— 不带标记
============  ================================================

角色 → 实际级别由分隔符决定（见 :meth:`LevelScheme.marker`）。
"""

from __future__ import annotations

import codecs
import re

# 角色常量
PARENT = "parent"
CHILD = "child"
DEEP = "deep"
BODY = "body"

# 默认分隔符：与 Dify 父子分段里「章节=一级标题」的约定一致。
# 写成字面量 ``\n# ``（反斜杠 n），用户直接复制分段节点上的值即可。
DEFAULT_PARENT_DELIMITER = "\\n# "
DEFAULT_CHILD_DELIMITER = "\\n## "

# Dify 侧父子分段插件的默认（无标题语义，纯按段落切）
_FALLBACK_PARENT = 1
_FALLBACK_CHILD = 2

# 标题级别上限（Markdown 最多 6 级）
MAX_HEADING_LEVEL = 6

_RE_HASH = re.compile(r"^#+")

# 角色 → 相对父级别的偏移量。
#   PARENT → 0（就是父分隔符对应的级别）
#   CHILD  → +1（子分隔符通常比父深一级）
#   DEEP   → +1 起步，再按逻辑深度递增
_ROLE_OFFSET = {PARENT: 0, CHILD: 1, DEEP: 1}


def decode_delimiter(delimiter: str) -> str:
    """按 Dify 的方式解码分隔符（``unicode_escape``）。

    Dify 侧：``codecs.decode(fixed_separator, "unicode_escape")``。
    非 ASCII 内容可能解码失败 —— 失败时原样返回（级别判定只关心 ``#``）。
    """
    if not delimiter:
        return ""
    try:
        return codecs.decode(delimiter, "unicode_escape")
    except (UnicodeDecodeError, ValueError):
        return delimiter


def parse_delimiter_level(delimiter: str) -> int | None:
    """分隔符 → 标题级别。

    ``\\n# `` → 1，``\\n## `` → 2，``\\n### `` → 3。
    分隔符里没有 ``#`` → 返回 ``None``（切分与标题标记无关）。
    """
    s = decode_delimiter(delimiter or "").lstrip()
    if not s.startswith("#"):
        return None
    m = _RE_HASH.match(s)
    return len(m.group(0)) if m else None


class LevelScheme:
    """按用户填的分隔符，推导各角色应当输出的标题级别。

    Parameters
    ----------
    parent_level:
        父分隔符对应的级别（``\\n# `` → 1）。``None`` 表示分隔符里没有
        ``#``，退回默认 1。
    child_level:
        子分隔符对应的级别（``\\n## `` → 2）。``None`` → 默认 2。

    级别被夹在 ``1..6`` 之间。若子级别不大于父级别（Dify 侧配置有误，
    或父子分隔符填反了），子级别被抬到「父 + 1」—— 否则两级分隔符
    会互相抢匹配，父子结构必然错乱。
    """

    def __init__(self, parent_level: int | None, child_level: int | None) -> None:
        p = parent_level if parent_level else _FALLBACK_PARENT
        c = child_level if child_level else _FALLBACK_CHILD
        p = max(1, min(p, MAX_HEADING_LEVEL))
        c = max(1, min(c, MAX_HEADING_LEVEL))
        if c <= p:
            c = min(p + 1, MAX_HEADING_LEVEL)
        self.parent_level = p
        self.child_level = c

    @classmethod
    def from_delimiters(cls, parent_delimiter: str, child_delimiter: str) -> "LevelScheme":
        """从两个分隔符构造。"""
        return cls(parse_delimiter_level(parent_delimiter),
                   parse_delimiter_level(child_delimiter))

    @classmethod
    def default(cls) -> "LevelScheme":
        """默认约定（父 1 级 / 子 2 级）—— 与插件历史行为一致。"""
        return cls(_FALLBACK_PARENT, _FALLBACK_CHILD)

    # ---- 角色 → 标记 ----

    def level_for(self, role: str, depth: int = 1) -> int:
        """角色 + 逻辑深度 → 实际标题级别。

        ``depth`` 只对 ``DEEP`` 有意义：``DEEP`` 的第一层是「比子块更深
        一层」，更深一层再深一级。``PARENT`` / ``CHILD`` 忽略它。
        """
        if role == BODY:
            return 0
        if role == PARENT:
            return self.parent_level
        if role == CHILD:
            return self.child_level
        # DEEP：child_level + depth（depth 从 1 起）
        return min(self.child_level + max(depth, 1), MAX_HEADING_LEVEL)

    def marker(self, role: str, depth: int = 1) -> str:
        """角色 → Markdown 标题标记（含尾随空格）。``BODY`` 返回空串。"""
        lv = self.level_for(role, depth)
        return ("#" * lv + " ") if lv else ""

    # ---- 便于排查 ----

    def describe(self) -> str:
        return (f"parent={self.parent_level}, child={self.child_level}")
