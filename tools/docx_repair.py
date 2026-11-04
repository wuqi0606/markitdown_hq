# -*- coding: utf-8 -*-
"""修复 docx 里指向不存在目标的关系引用。

问题
----
Word 文档里偶尔会残留**指向空目标的关系**：

    word/_rels/document.xml.rels
      <Relationship Id="rId48" Type=".../image" Target="../NULL"/>

正文里对应位置有一张图片（``<a:blip r:embed="rId48"/>``），
但目标文件并不在包里 —— 典型来源是「图片被删了，关系没清干净」。

mammoth 解析时会去 zip 里取 ``word/../NULL``，找不到就抛
``KeyError: There is no item named 'word/../NULL' in the archive``。
MarkItDown 把它包成 ``FileConversionException``，
**整个文件一条内容都产不出来** —— 一张坏图拖垮整篇文档。

实测
----
``项目交付操作指引.docx``：rId48 指向 ``../NULL``，
所在段落没有任何文字、图片尺寸高 0 EMU（残留空占位），
删掉后转换正常（14912 字符）。

处理方式
--------
只删**指向不存在目标**的关系，以及正文里对它们的引用：

* 关系项：从 ``.rels`` 里移除
* 正文引用：删掉整段 ``<w:drawing>``（DrawingML）
  与 ``<w:pict>``（VML）—— 顺带把只含坏图的空段落清掉

**不动**其他任何内容：文字、表格、正常图片、样式、编号全部原样保留。
"""

from __future__ import annotations

import io
import re
import zipfile

# 关系项：<Relationship Id=".." Type=".." Target=".."/>
_RE_REL = re.compile(r"<Relationship\b[^>]*/>")
# 属性：Id / Target
_RE_ID = re.compile(r'\bId="([^"]*)"')
_RE_TARGET = re.compile(r'\bTarget="([^"]*)"')
# 目标指向包外的（http:// 等）
_RE_EXTERNAL = re.compile(r"^(?:https?:|mailto:|file:|ftp:)", re.I)

# DrawingML 图片：<w:drawing>...</w:drawing>
_RE_DRAWING = re.compile(r"<w:drawing>.*?</w:drawing>", re.S)
# VML 图片：<w:pict>...</w:pict>
_RE_PICT = re.compile(r"<w:pict>.*?</w:pict>", re.S)
# 正文里对关系的引用
_RE_EMBED = re.compile(r'r:(?:embed|id|link)="([^"]*)"')


def _target_exists(names: set[str], base: str, target: str) -> bool:
    """判断关系目标在包里是否存在。

    ``base`` 是关系文件所在目录（如 ``word/_rels``），
    ``Target`` 是相对路径（相对关系文件所在目录的上一级，即部件所在目录）。
    """
    if not target or _RE_EXTERNAL.match(target):
        return True  # 外链不算缺失
    t = target.replace("\\", "/")
    # 关系文件的目录是 <part>/_rels，目标相对于 <part>
    part_dir = base.rsplit("/_rels", 1)[0] if "/_rels" in base else ""
    cand = f"{part_dir}/{t}" if part_dir else t
    # 归一化 . 与 ..
    parts: list[str] = []
    for seg in cand.split("/"):
        if seg in ("", "."):
            continue
        if seg == "..":
            if parts:
                parts.pop()
            continue
        parts.append(seg)
    norm = "/".join(parts)
    return norm in names


def find_broken_relations(blob: bytes) -> list[tuple[str, str]]:
    """找出指向不存在目标的关系，返回 [(rels 文件名, 关系 Id)]。"""
    broken: list[tuple[str, str]] = []
    try:
        zf = zipfile.ZipFile(io.BytesIO(blob))
    except zipfile.BadZipFile:
        return broken
    with zf:
        names = set(zf.namelist())
        for name in names:
            if not name.endswith(".rels"):
                continue
            try:
                data = zf.read(name).decode("utf-8", "replace")
            except KeyError:
                continue
            for m in _RE_REL.finditer(data):
                tag = m.group(0)
                idm = _RE_ID.search(tag)
                tm = _RE_TARGET.search(tag)
                if not idm or not tm:
                    continue
                if not _target_exists(names, name.rsplit("/", 1)[0] if "/" in name else "", tm.group(1)):
                    broken.append((name, idm.group(1)))
    return broken


def repair_docx(blob: bytes) -> tuple[bytes, int]:
    """修复 docx，返回 (修复后的字节, 移除的关系数)。

    没有坏关系时原样返回，字节完全一致。
    """
    broken = find_broken_relations(blob)
    if not broken:
        return blob, 0

    # 关系 Id → 所在 .rels 文件
    by_rels: dict[str, set[str]] = {}
    for rels_name, rid in broken:
        by_rels.setdefault(rels_name, set()).add(rid)
    all_bad_ids: set[str] = set()
    for ids in by_rels.values():
        all_bad_ids |= ids

    zin = zipfile.ZipFile(io.BytesIO(blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name in by_rels:
                text = data.decode("utf-8", "replace")

                def drop_rel(m: re.Match[str]) -> str:
                    tag = m.group(0)
                    idm = _RE_ID.search(tag)
                    return "" if idm and idm.group(1) in by_rels[name] else tag

                text = _RE_REL.sub(drop_rel, text)
                data = text.encode("utf-8")
            elif name == "word/document.xml":
                text = data.decode("utf-8", "replace")

                def drop_drawing(m: re.Match[str]) -> str:
                    seg = m.group(0)
                    ids = set(_RE_EMBED.findall(seg))
                    return "" if ids & all_bad_ids else seg

                text = _RE_DRAWING.sub(drop_drawing, text)
                text = _RE_PICT.sub(drop_drawing, text)
                data = text.encode("utf-8")
            zout.writestr(name, data)

    return out.getvalue(), len(broken)
