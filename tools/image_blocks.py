# -*- coding: utf-8 -*-
"""把「图片 + 描述」单独切成一个子块。

为什么
----
图片的签名 URL（``...jpg?timestamp=…&nonce=…&sign=…``）有 150+ 字符，
其中大半是随机串。混在正文里会稀释整个子块的语义：
实测「合同起草」所在子块里 URL 占了七成篇幅，召回时向量基本在描述
「一堆 URL 字符串」而不是「合同起草」。

把它单独切出去，正文子块就干净了。

怎么切
------
Dify 的父子分段是**纯文本分隔符切分**，插件无法直接命令 Dify
「把这几行放进哪个块」—— 唯一的手段是在文本里放一个分隔符。

所以做法是：在图片块前插一个**子分隔符级别的标记**，
让 Dify 切子块时正好切在图片前面。标记级别必须跟着用户填的
子分隔符走（``\\n## `` → 二级标记；``\\n### `` → 三级标记），
写死会导致换配置后失效。

插在图片所在的**更深处**（四级标题下的图片）也要用子分隔符级别：
Dify 只认分隔符字面量，不认标题的语义深浅。

为什么标记放在标题重建之后
--------------------------
插入的标记本身是一个标题。插件的层级判定靠编号形态
（形态序 / 点分体系），半路多出来的标记会被当成真标题参与判定，
把层级搞乱。所以这一步必须在所有标题重建完成之后做。

同级的标记会不会被 Dify 切错
--------------------------
父分隔符与子分隔符存在前缀包含关系（``\\n# `` 与 ``\\n## ``），
但只要标记用的是**子分隔符那个级别**，Dify 切父块时不会命中它
（``\\n# `` 匹配不上 ``\\n## X``）—— 父块边界不受影响。
"""

from __future__ import annotations

import re

from tools.heading_normalizer import _code_mask

# 插件插入的图片描述行：> 图片说明：xxx
_RE_DESC_LINE = re.compile(r"^>[ \t]*图片说明[：:].*$")

# markdown 图片引用（http/https 链接形态）
_RE_IMG_LINK = re.compile(r"^[ \t]*!\[[^\]]*\]\(\s*https?://[^)\s]+\s*\)[ \t]*$")

# 空行
_RE_BLANK = re.compile(r"^[ \t]*$")


def isolate_image_blocks(
    text: str,
    marker_line: str,
) -> tuple[str, int]:
    """把「图片描述 + 图片链接」包进一个标记行，独立成一个子块。

    参数
    ----
    text:
        待处理文本。
    marker_line:
        插在图片块前的标记行（如 ``"### 图片"``）。
        为空则不做任何处理。

    返回
    ----
    (新文本, 切出的图片块数)。
    """
    if not text or not marker_line:
        return text, 0

    lines = text.split("\n")
    in_code = _code_mask(lines)

    # 找出所有图片块 —— 「描述行 + 图片行」相邻算一块，
    # 单独一个图片行也算一块（「上传并替换为链接」模式下没有描述）。
    blocks: list[tuple[int, int]] = []  # (起始行号, 结束行号)
    i = 0
    n = len(lines)
    while i < n:
        if i in in_code:
            i += 1
            continue
        if _RE_DESC_LINE.match(lines[i]):
            # 描述行 —— 往下找紧邻的图片行（允许中间隔一个空行）
            j = i + 1
            if j < n and j not in in_code and _RE_BLANK.match(lines[j]):
                j += 1
            if j < n and j not in in_code and _RE_IMG_LINK.match(lines[j]):
                blocks.append((i, j))
                i = j + 1
                continue
            i += 1
            continue
        if _RE_IMG_LINK.match(lines[i]):
            blocks.append((i, i))
            i += 1
            continue
        i += 1

    if not blocks:
        return text, 0

    out: list[str] = []
    prev = 0
    for start, end in blocks:
        out.extend(lines[prev:start])
        out.append(marker_line)
        out.extend(lines[start:end + 1])
        prev = end + 1
    out.extend(lines[prev:])

    return "\n".join(out), len(blocks)
