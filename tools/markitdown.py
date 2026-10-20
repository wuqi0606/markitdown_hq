# -*- coding: utf-8 -*-
"""MarkItDown 文档转换 —— Dify Knowledge Pipeline 文档处理器。

包装微软 MarkItDown（0.1.8），并针对 RAG 场景做输出加工：

- 剥离base64 图片占位符（未开启 keep_images 时），避免撑爆 embedding 上下文
- 清理 Word 转换残留的空表格行（`| | | |`）
- 清理目录页残留
- **补齐丢失的标题层级**（restore_headings，默认开启）

关于最后一项：这批企业制度文件的章节标题在 Word 里用的是**普通加粗段落**
而非「标题1/2/3」样式，MarkItDown 靠样式映射生成 `#`，导致转换后层级全部丢失。
不补齐的话，父子分段的`\n#` / `\n##` 分隔符完全失效。
实测 31 份文件：补齐前 7 份零标题、10 份只有一级；补齐后 0 份无标题。

用法（作为 Dify 插件 Tool 调用）::

    invoke(files=[...], clean_noise=True, restore_headings=True,
           add_frontmatter=False, keep_images=False)
"""

from __future__ import annotations

import io
import re
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from markitdown import MarkItDown

from tools.heading_restorer import restore_headings
from tools.pdf_cleaner import clean_pdf_markdown

# ------------------------------------------------------------------ 清洗规则

# base64 图片：![](data:image/png;base64,....)
# 三个坑：
#   1) 覆盖 png/jpeg/gif/bmp/webp/x-emf/x-wmf 等全部 MIME，不能只列常见格式
#   2) 相邻图片无分隔符：](data:...)![alt](data:...)，必须非贪婪逐个匹配
#   3) alt 文本可能含 ']'（Word 图片名是乱码），不能写 \[[^\]]*\]
#      改为 [^	]*? 由后面的 ](data:image/ 来锚定结尾
_RE_IMAGE_PLACEHOLDER = re.compile(r"!\[[^\n]*?\]\(\s*data:image/[^)]*?\)")
# 兜底：清理没有被 markdown 图片语法包裹的裸 base64 片段
_RE_BARE_DATA_URI = re.compile(r"\(\s*data:image/[^)]*?\)")
# 仅含分隔符的空表格行：| | | |
_RE_EMPTY_ROW = re.compile(r"^\|(?:\s*\|)+\s*$")
# 目录条目（纯文本形态）：一、 编制目的 3 / （一） 进度管控 5
_RE_TOC_TEXT = re.compile(r"^\s*(?:[一二三四五六七八九十]+、|（[一二三四五六七八九十]+）)\s*\S+\s*\d+\s*$")
# 目录条目（markdown 链接形态）：[一、 编制目的 3](#_Toc23278)
_RE_TOC_LINK = re.compile(r"^\s*\[.+?\]\(#.+?\)\s*$")
# 目录区块内降级残留：Word 目录样式被 MarkItDown 降级为「标题名 + 页码」
#   例：「某某科技股份有限公司 1」「附件一、各业务线标准交付周期表 9」
# 仅在「目录」到下一个标题之间应用，避免误删正文里的数字结尾句子
_RE_TOC_TAIL_NUM = re.compile(r"^\s*\S[^\n]{0,44}?\s+\d{1,3}\s*$")
# 连续 3 个以上空行压缩为 2 个
_RE_MULTI_BLANK = re.compile(r"\n{3,}")

# 文件头魔数
_SIG_ZIP = b"PK"              # OOXML (.docx/.xlsx/.pptx) 本质是 zip
_SIG_OLE = b"\xd0\xcf\x11\xe0"  # Word 97-2003 / Excel 97-2003
_SIG_PDF = b"%PDF"


class MarkitdownTool(Tool):
    """将 PDF / Word / Excel / PPT / HTML / 图片 转换为干净 Markdown。"""

    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        files = tool_parameters.get("files") or []
        clean_noise = _as_bool(tool_parameters.get("clean_noise", True), True)
        add_frontmatter = _as_bool(tool_parameters.get("add_frontmatter", False), False)
        keep_images = _as_bool(tool_parameters.get("keep_images", False), False)
        fix_headings = _as_bool(tool_parameters.get("restore_headings", True), True)
        clean_pdf = _as_bool(tool_parameters.get("clean_pdf_headers", True), True)
        pdf_header_keywords = str(tool_parameters.get("pdf_header_keywords", "") or "").strip()

        if not files:
            yield self.create_text_message("No files provided.")
            return

        md = MarkItDown(enable_plugins=False)
        blocks: list[str] = []
        notices: list[str] = []

        for f in files:
            name = getattr(f, "filename", None) or getattr(f, "name", "unnamed")
            blob = _extract_blob(f)
            if blob is None:
                notices.append(f"[skip] {name}: unsupported file object")
                continue

            hint = _format_hint(blob)
            if hint == "legacy_doc":
                notices.append(
                    f"[skip] {name}: legacy Word 97-2003 (.doc) detected "
                    f"despite .docx extension — MarkItDown cannot parse it. "
                    f"Re-save as real .docx first."
                )
                continue

            try:
                # convert_stream 需要 file-like 对象，不能直接传 bytes
                stream = io.BytesIO(blob)
                result = md.convert_stream(stream, file_extension=_ext_of(name, hint))
                raw = result.markdown or ""
            except Exception as exc:  # noqa: BLE001
                notices.append(f"[fail] {name}: {type(exc).__name__}: {exc}")
                continue

            text = _clean(raw, clean_noise=clean_noise, keep_images=keep_images)

            # PDF 专用清洗：页眉/页脚/页码。必须放在通用清洗之后 ——
            # 通用清洗已删掉空表格行，页码的「距页眉距离」判定依赖这些行做间隔。
            if clean_pdf and hint == "pdf":
                text, pstats = clean_pdf_markdown(
                    text, header_keywords=pdf_header_keywords
                )
                if sum(pstats.values()):
                    notices.append(
                        f"[info] {name}: pdf cleanup - "
                        f"headers {pstats['headers_removed']}, "
                        f"page-numbers {pstats['page_nums_removed']}, "
                        f"table-sep {pstats['sep_rows_removed']}"
                    )

            if fix_headings:
                # 标题补齐必须在清洗之后 —— 清洗会把「一、 xxx 3」这类目录条目删掉，
                # 先补标题的话目录残留会被误判成真标题
                text, added = restore_headings(text)
                if added:
                    notices.append(f"[info] {name}: restored {added} heading(s)")
            if add_frontmatter:
                text = f"---\nsource: {name}\n---\n\n{text}"
            blocks.append(text)

        output = "\n\n---\n\n".join(blocks)

        if notices:
            output = "<!-- conversion notices\n" + "\n".join(notices) + "\n-->\n\n" + output

        yield self.create_text_message(output)


# ------------------------------------------------------------------ 辅助函数


def _as_bool(value: Any, default: bool) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"true", "1", "yes", "on"}
    return bool(value)


def _extract_blob(f: Any) -> bytes | None:
    """从 Dify 的各种文件对象形态里取出字节内容。"""
    for attr in ("blob", "content", "data"):
        v = getattr(f, attr, None)
        if isinstance(v, (bytes, bytearray)):
            return bytes(v)
    inner = getattr(f, "file", None)
    if inner is not None:
        return _extract_blob(inner)
    if isinstance(f, (bytes, bytearray)):
        return bytes(f)
    return None


def _format_hint(blob: bytes) -> str:
    if blob.startswith(_SIG_ZIP):
        return "zip"
    if blob.startswith(_SIG_OLE):
        return "legacy_doc"
    if blob.startswith(_SIG_PDF):
        return "pdf"
    return "unknown"


def _ext_of(name: str, hint: str) -> str:
    """MarkItDown 需要扩展名来决定用哪个 converter。"""
    tail = name.rsplit(".", 1)
    if len(tail) == 2 and len(tail[1]) <= 6:
        return "." + tail[1].lower()
    return {"zip": ".docx", "pdf": ".pdf", "legacy_doc": ".doc"}.get(hint, ".txt")


_RE_TABLE_SEP = re.compile(r"^\|[\s:\-|]+\|$")


def _strip_images(text: str) -> str:
    """移除 base64 图片引用。

    两轮：先删完整的 markdown 图片语法，再兜底删裸的 data:image URI。
    第二轮是必要的 —— Word 的图片名常含乱码字符（如 @YQNWY_ST]A`N4WA$I$[6US），
    alt 文本里会带 ']'，导致第一轮正则锚定失败而漏删。
    """
    text = _RE_IMAGE_PLACEHOLDER.sub("", text)
    text = _RE_BARE_DATA_URI.sub("", text)
    return text


def _fix_table_sep_position(text: str) -> tuple[str, int]:
    """把顶到表格第一行的分隔行移回真表头之后。

    markitdown 输出 Word 表格的顺序是：空表头行 → 分隔行 → 真表头 → 数据。
    前一步删掉空表头行后变成：分隔行 → 真表头 → 数据，分隔行成了表格
    第一行，渲染时真表头被降级为普通数据行，表格失去表头。

    判定：某行是分隔行、且它上方不是表格行（即它是表格块首行）、
    且下方紧跟以 | 开头的非分隔行 —— 把分隔行与下一行交换。
    返回 (修复后文本, 修复表格数)。
    """
    lines = text.split("\n")
    out: list[str] = []
    fixed = 0
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        is_sep = bool(_RE_TABLE_SEP.match(line.strip()))
        prev_is_table = bool(out) and out[-1].lstrip().startswith("|")
        if is_sep and not prev_is_table and i + 1 < n:
            nxt = lines[i + 1]
            if nxt.strip().startswith("|") and not _RE_TABLE_SEP.match(nxt.strip()):
                out.append(nxt)
                out.append(line)
                fixed += 1
                i += 2
                continue
        out.append(line)
        i += 1
    return "\n".join(out), fixed


def _clean(raw: str, *, clean_noise: bool, keep_images: bool) -> str:
    text = raw or ""
    if not clean_noise:
        if not keep_images:
            text = _strip_images(text)
        return text.strip() + "\n"

    # 1. base64 图片占位符 —— 会撑爆 embedding 上下文
    if not keep_images:
        text = _strip_images(text)

    # 2. 空表格行 —— 纯噪声
    #    实测 markitdown 输出顺序：空表头行 | | |  →  | --- |  →  真表头 → 数据。
    #    删掉空表头行后分隔行会顶到表格第一行，真表头被降级为数据行。
    #    修复见 _fix_table_sep_position，在删行之后统一处理。
    text = "\n".join(
        line for line in text.split("\n")
        if not (line.strip() and _RE_EMPTY_ROW.match(line))
    )
    text, tables_fixed = _fix_table_sep_position(text)

    # 3. 目录残留 —— 两种形态：
    #    (a) 带编号的「一、 xxx 3」/「[标题](#锚点)」，全文范围内清理
    #    (b) 降级残留「标题名 3」，只在「目录」到下一个标题之间清理
    out: list[str] = []
    in_toc = False
    for line in text.split("\n"):
        if _RE_TOC_TEXT.match(line) or _RE_TOC_LINK.match(line):
            continue
        if line.strip() == "目录":
            in_toc = True
            continue
        if in_toc:
            if line.startswith("#"):
                in_toc = False
            elif line.strip() and _RE_TOC_TAIL_NUM.match(line):
                continue
        out.append(line)
    text = "\n".join(out)

    # 4. 压缩连续空行
    text = _RE_MULTI_BLANK.sub("\n\n", text)

    return text.strip() + "\n"