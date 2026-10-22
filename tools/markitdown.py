# -*- coding: utf-8 -*-
"""MarkItDown 文档转换 —— Dify Knowledge Pipeline 文档处理器。

包装微软 MarkItDown（0.1.8），并针对 RAG 场景做输出加工：

- 图片处理三选一（image_mode）：remove 删除 base64 引用（默认，避免撑爆
  embedding 上下文）；url 通过 SDK 反向调用上传到 Dify 存储并替换为可访问链接
- 清理 Word 转换残留的空表格行（`| | | |`），并修复分隔行错位
- 清理目录页残留
- **补齐丢失的标题层级**（restore_headings，默认开启）：Word 里未套用
  标题样式、仅以加粗呈现的章节标题，MarkItDown 转换后层级全丢，
  不补齐则父子分段的 `\n#` / `\n##` 分隔符完全失效。
  层级按文档实际出现的编号形态动态映射。

用法（作为 Dify 插件 Tool 调用）::

    invoke(files=[...], clean_noise=True, restore_headings=True,
           add_frontmatter=False, image_mode='remove')
"""

from __future__ import annotations

import base64
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
        fix_headings = _as_bool(tool_parameters.get("restore_headings", True), True)
        clean_pdf = _as_bool(tool_parameters.get("clean_pdf_headers", True), True)
        pdf_header_keywords = str(tool_parameters.get("pdf_header_keywords", "") or "").strip()
        image_mode = str(tool_parameters.get("image_mode", "remove") or "remove").strip().lower()
        if image_mode not in {"remove", "url", "describe"}:
            image_mode = "remove"

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
                # convert_stream 需要 file-like 对象，不能直接传 bytes。
                # keep_data_uris：markitdown 的 html→md 阶段默认把 <img> 的
                # data URI 截断成 "data:image/png;base64..." 占位符（三点是
                # 字面量，不是真实数据）。url 模式必须开此开关才能拿到
                # 真实图片字节，供后续上传换 URL。
                stream = io.BytesIO(blob)
                result = md.convert_stream(
                    stream,
                    file_extension=_ext_of(name, hint),
                    keep_data_uris=(image_mode == "url"),
                )
                raw = result.markdown or ""
            except Exception as exc:  # noqa: BLE001
                notices.append(f"[fail] {name}: {type(exc).__name__}: {exc}")
                continue

            # 图片处理：remove=删掉 base64 引用；url=上传换签名链接；
            # describe=上传后再逐张调视觉模型生成描述，插在 URL 前。
            # 上传/描述必须在这里做（需要 self.session），不能放进纯文本的 _clean。
            if image_mode in {"url", "describe"}:
                raw, n_up, n_fail, uploaded = _upload_images(
                    raw, self.session, name, notices
                )
                if n_up:
                    notices.append(f"[info] {name}: images uploaded {n_up}, failed {n_fail}")
                raw = _strip_images(raw)  # 清掉上传失败残留的 data URI

                if image_mode == "describe":
                    # model 参数为字符串「提供商ID/模型名」（model-selector
                    # 类型在知识库流水线节点面板不受支持，会导致前端崩溃）。
                    # rpartition('/')：提供商 ID 本身含斜杠（如
                    # langgenius/volcengine_maas），取最后一个 / 分割。
                    raw_model = str(tool_parameters.get("model", "") or "").strip()
                    provider, _sep, model_name = raw_model.rpartition("/")
                    model_config = None
                    if provider and model_name:
                        model_config = {
                            "provider": provider,
                            "model": model_name,
                            "mode": "chat",
                            "completion_params": {},
                        }
                    if not model_config:
                        notices.append(
                            f"[warn] {name}: 图片处理方式为生成描述，但未正确填写"
                            f"「视觉模型」（格式：提供商ID/模型名），跳过描述生成"
                        )
                    else:
                        raw, n_d, n_df = _describe_images(
                            raw, self.session, model_config, name, notices,
                            uploaded=uploaded,
                        )
                        if n_d or n_df:
                            notices.append(
                                f"[info] {name}: image descriptions generated {n_d}, failed {n_df}"
                            )

            text = _clean(raw, clean_noise=clean_noise)

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


# data URI 完整形态：![alt](data:image/png;base64,AAAA)
# base64 字符集不含 ')'，用 [^)] 安全截止；alt 可能含乱码，单独兜底删裸 URI
_RE_MD_DATA_IMAGE = re.compile(
    r"!\[[^\]]*\]\(\s*data:image/([^;)]+);base64,([^)\s]+)\s*\)"
)


def _upload_images(
    text: str, session, source_name: str, notices: list[str]
) -> tuple[str, int, int, list[dict]]:
    """把文本里的 base64 图片上传到 Dify 存储，替换为签名 URL。

    使用 SDK 的反向调用能力（session.file.upload），图片二进制
    存入 Dify 文件存储后拿到 preview_url（签名链接，有时效）。
    单张失败只删该图并在 notices 记录，不影响其余图片。
    返回 (替换后文本, 成功数, 失败数, uploaded 清单)。
    uploaded 清单元素：{"url": 签名链接, "b64": 原始 base64 字符串,
    "mime": image/xxx}——b64 供 describe 模式直接内嵌模型请求，
    免去按 URL 回源下载（签名 URL 是内网地址，模型供应商拉不到）。
    """
    if "data:image/" not in text:
        return text, 0, 0, []

    from dify_plugin.invocations.file import File

    file_invocation = File(session)
    out: list[str] = []
    uploaded: list[dict] = []
    ok = fail = 0
    last = 0
    for i, m in enumerate(_RE_MD_DATA_IMAGE.finditer(text), start=1):
        out.append(text[last : m.start()])
        last = m.end()
        ext = m.group(1).lower()
        b64_data = m.group(2)
        try:
            data = base64.b64decode(b64_data)
            resp = file_invocation.upload(
                filename=f"{source_name}-image-{i}.{ext}",
                content=data,
                mimetype=f"image/{ext}",
            )
            url = (getattr(resp, "preview_url", None) or "").strip()
            if url:
                out.append(f"![image]({url})")
                uploaded.append(
                    {"url": url, "b64": b64_data, "mime": f"image/{ext}", "ext": ext}
                )
                ok += 1
            else:
                out.append("")
                fail += 1
                notices.append(f"[warn] {source_name}: image {i} uploaded but no url")
        except Exception as exc:  # noqa: BLE001
            out.append("")
            fail += 1
            notices.append(f"[warn] {source_name}: image {i} upload failed: {exc}")
    out.append(text[last:])
    return "".join(out), ok, fail, uploaded


_DESCRIBE_PROMPT = (
    "用一句中文描述这张图片：它是什么类型（流程图/审批图/界面截图/组织架构图等），"
    "包含哪些关键节点、步骤或字段。直接输出描述本身，不要任何前后缀或引号。"
)


def _fetch_image_b64(session, url: str) -> str | None:
    """下载图片 URL 转 base64。用于 MinerU 场景——文本里只有 URL 没有字节。

    关键：模型供应商（如豆包云端）拉不到内网签名 URL（10.x 私网地址），
    所以图片必须以 base64 内嵌进模型请求，由插件进程（与 Dify 同网络）
    负责取回字节。下载失败返回 None。
    """
    try:
        import requests

        resp = requests.get(url, timeout=60)
        if resp.status_code != 200:
            return None
        return base64.b64encode(resp.content).decode()
    except Exception:  # noqa: BLE001
        return None


def _describe_images(
    text: str,
    session,
    model_config: dict,
    source_name: str,
    notices: list[str],
    uploaded: list[dict] | None = None,
) -> tuple[str, int, int]:
    """对文本里的图片链接逐张调视觉模型生成描述，插在引用之前。

    使用 SDK 的 LLM 反向调用（session.model.llm.invoke），模型用平台
    上已配置的（model_selector 参数选的），不需要插件管凭据。
    图片数据一律以 base64 内嵌进模型请求（不从 URL 让模型拉取——
    签名 URL 是内网地址，云端模型供应商访问不到）：
    - docx 场景：uploaded 清单里带原始 base64（上传时留下的），直接用
    - MinerU 场景：只有 URL，由插件进程下载转 base64（含签名，可访问）
    单张失败只跳过该图并记 notice。返回 (新文本, 成功数, 失败数)。
    """
    url_re = re.compile(r"!\[([^\]]*)\]\((https?://[^)]+)\)")
    targets = url_re.findall(text)
    if not targets:
        return text, 0, 0

    b64_by_url = {u["url"]: u["b64"] for u in (uploaded or [])}

    from dify_plugin.entities.model.message import (
        ImagePromptMessageContent,
        TextPromptMessageContent,
        UserPromptMessage,
    )
    from dify_plugin.invocations.model.llm import LLMInvocation

    llm = LLMInvocation(session)
    ok = fail = 0
    for i, (_alt, url) in enumerate(targets, start=1):
        try:
            b64 = b64_by_url.get(url)
            if b64 is None:
                b64 = _fetch_image_b64(session, url)
                if b64 is None:
                    fail += 1
                    notices.append(
                        f"[warn] {source_name}: image {i} download failed (signed url not reachable)"
                    )
                    continue
            ext = url.rsplit(".", 1)[-1].lower() if "." in url.rsplit("/", 1)[-1] else "jpg"
            mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(ext, "image/jpeg")
            prompt = UserPromptMessage(
                content=[
                    TextPromptMessageContent(data=_DESCRIBE_PROMPT),
                    ImagePromptMessageContent(
                        base64_data=b64, mime_type=mime, format=ext, filename=f"image-{i}"
                    ),
                ]
            )
            resp = llm.invoke(model_config=model_config, prompt_messages=[prompt], stream=False)
            desc = str(resp.message.content or "").strip().strip('"')
            if desc:
                text = text.replace(f"![{_alt}]({url})", f"> 图片说明：{desc}\n![{_alt}]({url})", 1)
                ok += 1
            else:
                fail += 1
                notices.append(f"[warn] {source_name}: image {i} empty description")
        except Exception as exc:  # noqa: BLE001
            fail += 1
            notices.append(f"[warn] {source_name}: image {i} describe failed: {exc}")
    return text, ok, fail


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


def _clean(raw: str, *, clean_noise: bool) -> str:
    text = raw or ""
    if not clean_noise:
        # base64 图片已在 _invoke 阶段处理（删除或换 URL），这里只兜底残留
        text = _strip_images(text)
        return text.strip() + "\n"

    # 1. base64 图片占位符 —— 会撑爆 embedding 上下文。
    #    url 模式下 data URI 已在 _invoke 里替换成链接，这里删不到，无害。
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