# -*- coding: utf-8 -*-
"""docx 段落级元数据提取（直读 OOXML）。

为什么必须读原始 docx：
MarkItDown（内部走 mammoth）转 docx 时丢掉了三类结构信息 ——

1. **自动编号体系**（numId / ilvl / 编号格式 `%1、` `（%1）` `%1.` …）
   mammoth 会把不同 numId 的列表**合并重排**成连续序号，于是
   「1. 发起申请 / 2. 提前实施…」这种「标题与内容混在一列」就出现了，
   而原始 docx 里这两组根本是不同的编号体系。
   编号还可以从 `numbering.xml` 还原成本来的样子。

2. **字号与字体**
   手打编号的章节标题，靠「16 磅黑体」和「14 磅仿宋」区分；
   转成 markdown 后这两种完全看不出区别。

3. **段落样式（outlineLvl）**
   Word 标题层级只有一部分段落带，mammoth 认的是样式名不是 outlineLvl，
   所以带 outlineLvl 但样式名不标准的段落也会丢掉层级。

产出 `DocxMeta`：每个段落一条元数据，再由 `heading_restorer` 按文本对齐回
markdown 行（实测对齐率 96%）。
"""

from __future__ import annotations

import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from io import BytesIO
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 编号形态（从粗到细），决定了层级顺序
FORM_CHAPTER = "chapter"  # 第x章
FORM_CN = "cn"      # 一、二、  /  附件
FORM_ARTICLE = "article"  # 第x条
FORM_CNP = "cnp"    # （一）（二）
FORM_NUM = "num"    # 1、 2、 3、
FORM_NUMP = "nump"  # （1）（2）
FORM_ORDER = (FORM_CHAPTER, FORM_CN, FORM_ARTICLE, FORM_CNP, FORM_NUM, FORM_NUMP)
FORM_RANK = {f: i for i, f in enumerate(FORM_ORDER)}

# 文本形态识别（用于手打编号的段落）
RE_TEXT_FORM: tuple[tuple[str, re.Pattern[str]], ...] = (
    (FORM_CHAPTER, re.compile(r"^\s*第[一二三四五六七八九十百]+\s*[章节部分]")),
    (FORM_CN, re.compile(r"^\s*[一二三四五六七八九十]+\s*[、.．]")),
    (FORM_ARTICLE, re.compile(r"^\s*第[一二三四五六七八九十百]+\s*条")),
    (FORM_CNP, re.compile(r"^\s*[（(]\s*[一二三四五六七八九十]+\s*[)）]")),
    (FORM_NUM, re.compile(r"^\s*\d{1,2}\s*[、.．]")),
    (FORM_NUMP, re.compile(r"^\s*[（(]\s*\d{1,2}\s*[)）]")),
)
# 附件与「第x章」都算章级
RE_ATTACH = re.compile(r"^\s*附\s*件\s*[一二三四五六七八九十\d]+")

# 常见中文标题字体（文档里没有真标题样式时的兜底）
STATIC_HEAD_FONTS = frozenset({
    "黑体", "楷体_GB2312", "方正黑体", "方正黑体_GBK",
    "方正小标宋", "方正小标宋简体", "微软雅黑", "华文中宋", "华文细黑",
})

# 目录标题行
RE_TOC_MARK = re.compile(r"^\s*目\s*录\s*[:：]?\s*$")

# mammoth 靠样式名（heading N / 标题 N）判断是不是标题
RE_MD_HEADING_NAME = re.compile(r"^(?:heading|标题)\s*(\d)$", re.IGNORECASE)

# Word 里 outlineLvl=9 表示「正文级」，不是标题
OUTLINE_BODY = 9

# 句子结束标点 —— 标题不以这些结尾
SENT_END = set("。；！？，、；")

_CN_DIGITS = "零一二三四五六七八九"


def canonical_form(form: str | None) -> str | None:
    """归一化形态别名。"""
    if form is None:
        return None
    if form == "chapter":
        return FORM_CHAPTER
    if form == "article":
        return FORM_ARTICLE
    if form == "attach":
        return FORM_CN
    return form


def form_from_numbering(num_fmt: str | None, lvl_text: str | None) -> str | None:
    """由 (numFmt, lvlText) 推出编号形态。

    - `%1、` + chineseCounting → 一、二、   → cn
    - `%1、` + decimal         → 1、2、      → num
    - `（%1）` + chineseCounting → （一）   → cnp
    - `（%1）` + decimal         → （1）    → nump
    - `第%1章`                  → 第一章    → chapter
    - `第%1条`                  → 第一条    → article
    - bullet                    → 无编号语义 → None
    """
    if not lvl_text:
        return None
    lt = lvl_text.strip()
    if not lt or num_fmt == "bullet":
        return None
    if "章" in lt or "部分" in lt:
        return FORM_CHAPTER
    if "条" in lt:
        return FORM_ARTICLE
    paren = ("（" in lt) or ("(" in lt)
    if num_fmt in ("chineseCounting", "chineseCountingThousand", "chineseLegalSimplified"):
        return FORM_CNP if paren else FORM_CN
    if num_fmt in ("decimal", "decimalZero", "arabicAlpha", "arabicAbjad"):
        return FORM_NUMP if paren else FORM_NUM
    if num_fmt in ("lowerLetter", "upperLetter", "lowerRoman", "upperRoman"):
        return FORM_NUMP if paren else FORM_NUM
    return None


def form_from_text(text: str) -> str | None:
    """由段落文本推出编号形态（手打编号的段落用）。"""
    t = text.strip()
    if not t:
        return None
    if RE_ATTACH.match(t):
        return "attach"
    for key, pat in RE_TEXT_FORM:
        if pat.match(t):
            return key
    return None


def _cn_number(n: int) -> str:
    """中文数字（支持到 99，够用）。"""
    if n < 10:
        return _CN_DIGITS[n]
    if n < 20:
        return "十" + (_CN_DIGITS[n - 10] if n % 10 else "")
    t, o = divmod(n, 10)
    return _CN_DIGITS[t] + "十" + (_CN_DIGITS[o] if o else "")


def render_number(n: int, num_fmt: str | None) -> str:
    """把计数器渲染成该编号格式对应的数字。"""
    if num_fmt in ("chineseCounting", "chineseCountingThousand", "chineseLegalSimplified"):
        return _cn_number(n)
    if num_fmt == "lowerLetter" and 1 <= n <= 26:
        return chr(ord("a") + n - 1)
    if num_fmt == "upperLetter" and 1 <= n <= 26:
        return chr(ord("A") + n - 1)
    return str(n)


def render_lvl_text(lvl_text: str, counters: dict[int, int],
                    num_fmt: str | None, fallback: int) -> str:
    """按 `lvlText` 模板还原原始编号，如 `%1、` → `3、`、`（%1）` → `（二）`。

    `%N` 指第 N 级（1 起）的计数器。
    """
    def rep(m: re.Match[str]) -> str:
        lv = int(m.group(1)) - 1
        c = counters.get(lv, fallback)
        return render_number(c, num_fmt)

    return re.sub(r"%(\d)", rep, lvl_text)


@dataclass
class ParaMeta:
    """一个 docx 段落的元数据。"""

    index: int
    text: str
    style: str | None = None
    md_level: int | None = None          # mammoth 会渲染成几个 `#`（1 起）
    outline: int | None = None           # 作者的大纲级别（0 起，9=正文不算）
    para_outline: int | None = None      # 段级直接设的大纲级别（0 起）
    style_outline: int | None = None     # 样式表里带的大纲级别（0 起）
    num_id: str | None = None           # 自动编号体系 id
    ilvl: int | None = None             # 编号体系内的层级
    num_fmt: str | None = None
    lvl_text: str | None = None
    size: float | None = None           # 字号（磅，含样式回落）
    font: str | None = None             # 中文字体（含样式回落）
    run_font: str | None = None         # run 上显式设的字体（没有回落）
    bold: bool = False
    in_table: bool = False
    in_toc: bool = False                # 目录区块内的条目
    in_cover: bool = False              # 封面/版本记录等前置区
    ordinal: int = 0                   # 在 (numId, ilvl) 组内的序号（1 起）
    counters: dict[int, int] = field(default_factory=dict)  # 该段落时各级的计数快照
    form: str | None = None             # 编号形态（已 canonical 化）

    @property
    def is_real_heading(self) -> bool:
        """mammoth 会渲染成 `#` 的段落 —— 已经是 Markdown 标题。"""
        return self.md_level is not None

    @property
    def is_outlined(self) -> bool:
        """作者标了大纲级别，但 mammoth 未必认得（样式名不是 heading N）。"""
        return self.outline is not None

    @property
    def is_auto_numbered(self) -> bool:
        """Word 自动编号段落 —— mammoth 转成了列表项。"""
        return self.num_id is not None

    @property
    def key(self) -> tuple[str, str]:
        """自动编号分组键。"""
        return (self.num_id or "", str(self.ilvl or 0))

    @property
    def rank(self) -> int | None:
        """编号形态的粗细序号（越小越粗）。"""
        if self.form is None:
            return None
        return FORM_RANK.get(self.form)


@dataclass
class DocxMeta:
    """整篇文档的元数据。"""

    paras: list[ParaMeta] = field(default_factory=list)
    base_size: float | None = None          # 正文基准字号
    head_fonts: frozenset[str] = frozenset()  # 文档里当作标题用的字体
    heading_groups: frozenset[tuple[str, str]] = frozenset()  # 判定为标题组的编号组
    group_form: dict[tuple[str, str], str] = field(default_factory=dict)  # 编号组 → 形态
    num_counters: dict[str, dict[int, int]] = field(default_factory=dict)  # numId → 各级计数器
    body_start: int = 0                     # 正文主体起点（之前的属于封面/前置区）

    def by_index(self) -> dict[int, ParaMeta]:
        return {p.index: p for p in self.paras}


def _para_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def _run_format(p: ET.Element) -> tuple[float | None, str | None, bool]:
    """取第一个非空 run 的字号 / 字体 / 加粗（run 级，不做回落）。"""
    size: float | None = None
    font: str | None = None
    bold = False
    for r in p.findall(f"{W}r"):
        if not "".join(x.text or "" for x in r.iter(f"{W}t")).strip():
            continue
        rp = r.find(f"{W}rPr")
        if rp is not None:
            sz = rp.find(f"{W}sz")
            if sz is not None:
                try:
                    size = int(sz.get(f"{W}val")) / 2
                except (TypeError, ValueError):
                    size = None
            rf = rp.find(f"{W}rFonts")
            if rf is not None:
                font = rf.get(f"{W}eastAsia") or rf.get(f"{W}ascii")
            b = rp.find(f"{W}b")
            # <w:b w:val="0"/> 是显式取消加粗
            if b is not None and b.get(f"{W}val") not in ("0", "false"):
                bold = True
        break
    return size, font, bold


def _style_formats(styles: ET.Element | None) -> dict[str, tuple[str | None, float | None]]:
    """styleId → (中文字体, 字号)。取样式自己的 rPr 定义。"""
    out: dict[str, tuple[str | None, float | None]] = {}
    if styles is None:
        return out
    for st in styles.iter(f"{W}style"):
        sid = st.get(f"{W}styleId")
        if sid is None:
            continue
        rp = st.find(f"{W}rPr")
        if rp is None:
            continue
        font: str | None = None
        size: float | None = None
        rf = rp.find(f"{W}rFonts")
        if rf is not None:
            font = rf.get(f"{W}eastAsia") or rf.get(f"{W}ascii")
        sz = rp.find(f"{W}sz")
        if sz is not None:
            try:
                size = int(sz.get(f"{W}val")) / 2
            except (TypeError, ValueError):
                size = None
        if font or size:
            out[sid] = (font, size)
    return out


def _default_format(styles: ET.Element | None) -> tuple[str | None, float | None]:
    """docDefaults 里的字体 / 字号。"""
    if styles is None:
        return None, None
    dd = styles.find(f"{W}docDefaults")
    if dd is None:
        return None, None
    rp = dd.find(f"{W}rPrDefault/{W}rPr")
    if rp is None:
        return None, None
    font = size = None
    rf = rp.find(f"{W}rFonts")
    if rf is not None:
        font = rf.get(f"{W}eastAsia") or rf.get(f"{W}ascii")
    sz = rp.find(f"{W}sz")
    if sz is not None:
        try:
            size = int(sz.get(f"{W}val")) / 2
        except (TypeError, ValueError):
            size = None
    return font, size


def _para_format(p: ET.Element, style: str | None,
                 style_fmt: dict[str, tuple[str | None, float | None]],
                 default_fmt: tuple[str | None, float | None],
                 ) -> tuple[float | None, str | None, str | None, bool]:
    """段落的字号 / 字体 / run 字体 / 加粗。

    字号与字体按 **run 级 → 段落样式级 → 文档默认级** 回落 ——
    标题段落常把字体写在样式里（heading 1 → 黑体），run 上是空的；
    反过来正文段落继承 Normal（仿宋 16 磅），run 上也是空的。
    只靠 run 级信号区分不出两者。
    """
    size, font, bold = _run_format(p)
    run_font = font
    sf, ss = style_fmt.get(style, (None, None)) if style else (None, None)
    if font is None:
        font = sf
    if size is None:
        size = ss
    df, ds = default_fmt
    if font is None:
        font = df
    if size is None:
        size = ds
    return size, font, run_font, bold


def parse_docx_meta(blob: bytes) -> DocxMeta | None:
    """解析 docx 字节流；不是 docx 或解析失败返回 None。"""
    try:
        with zipfile.ZipFile(BytesIO(blob)) as zf:
            names = set(zf.namelist())
            if "word/document.xml" not in names:
                return None
            doc = ET.fromstring(zf.read("word/document.xml"))
            try:
                styles = ET.fromstring(zf.read("word/styles.xml"))
            except KeyError:
                styles = None
            try:
                numb = ET.fromstring(zf.read("word/numbering.xml"))
            except KeyError:
                numb = None
    except (zipfile.BadZipFile, ET.ParseError, KeyError, OSError):
        return None

    # styleId -> outlineLvl ；styleId -> mammoth 认得的标题层级
    outline: dict[str, int] = {}
    md_levels: dict[str, int] = {}
    if styles is not None:
        for st in styles.iter(f"{W}style"):
            sid = st.get(f"{W}styleId")
            if sid is None:
                continue
            o = st.find(f"{W}pPr/{W}outlineLvl")
            if o is not None:
                try:
                    lv = int(o.get(f"{W}val"))
                except (TypeError, ValueError):
                    lv = None
                if lv is not None and lv < OUTLINE_BODY:
                    outline[sid] = lv
            nm = st.find(f"{W}name")
            if nm is not None:
                m = RE_MD_HEADING_NAME.match((nm.get(f"{W}val") or "").strip())
                if m:
                    md_levels[sid] = int(m.group(1))

    # 样式级字体 / 字号（标题样式 → 黑体/楷体，正文样式 → 仿宋/宋体）
    style_fmt = _style_formats(styles)
    default_fmt = _default_format(styles)
    # 没有 pStyle 的段落继承「默认段落样式」（Word 里通常是 Normal）
    for st in (styles.iter(f"{W}style") if styles is not None else ()):
        if st.get(f"{W}type") == "paragraph" and st.get(f"{W}default") in ("1", "true"):
            sid = st.get(f"{W}styleId")
            if sid and sid in style_fmt:
                default_fmt = style_fmt[sid]
            break

    # numId -> abstractNumId ; (abstractId, ilvl) -> (numFmt, lvlText)
    num2abs: dict[str, str] = {}
    lvlmap: dict[tuple[str, str], tuple[str | None, str | None]] = {}
    if numb is not None:
        for n in numb.iter(f"{W}num"):
            a = n.find(f"{W}abstractNumId")
            nid = n.get(f"{W}numId")
            if nid is not None and a is not None:
                num2abs[nid] = a.get(f"{W}val") or ""
        for an in numb.iter(f"{W}abstractNum"):
            aid = an.get(f"{W}abstractNumId") or ""
            for lvl in an.findall(f"{W}lvl"):
                il = lvl.get(f"{W}ilvl") or "0"
                f = lvl.find(f"{W}numFmt")
                t = lvl.find(f"{W}lvlText")
                lvlmap[(aid, il)] = (
                    f.get(f"{W}val") if f is not None else None,
                    t.get(f"{W}val") if t is not None else None,
                )

    # 表格内段落
    in_tbl: set[int] = set()
    for tbl in doc.iter(f"{W}tbl"):
        for p in tbl.iter(f"{W}p"):
            in_tbl.add(id(p))

    paras: list[ParaMeta] = []
    for idx, p in enumerate(doc.iter(f"{W}p")):
        ppr = p.find(f"{W}pPr")
        style = num_id = ilvl = None
        para_outline = None
        if ppr is not None:
            ps = ppr.find(f"{W}pStyle")
            if ps is not None:
                style = ps.get(f"{W}val")
            np = ppr.find(f"{W}numPr")
            if np is not None:
                e = np.find(f"{W}ilvl")
                ilvl = e.get(f"{W}val") if e is not None else "0"
                e = np.find(f"{W}numId")
                num_id = e.get(f"{W}val") if e is not None else None
                if num_id == "0":        # numId=0 表示「取消编号」
                    num_id = None
                    ilvl = None
            oe = ppr.find(f"{W}outlineLvl")
            if oe is not None:
                try:
                    lv = int(oe.get(f"{W}val"))
                except (TypeError, ValueError):
                    lv = None
                if lv is not None and lv < OUTLINE_BODY:
                    para_outline = lv
        text = _para_text(p)
        size, font, run_font, bold = _para_format(p, style, style_fmt, default_fmt)
        num_fmt = lvl_text = None
        if num_id is not None:
            num_fmt, lvl_text = lvlmap.get((num2abs.get(num_id, ""), str(ilvl or "0")),
                                           (None, None))
        if num_id is not None:
            form = canonical_form(form_from_numbering(num_fmt, lvl_text))
        else:
            form = canonical_form(form_from_text(text))
        style_outline = outline.get(style) if style else None
        paras.append(ParaMeta(
            index=idx, text=text, style=style,
            md_level=md_levels.get(style) if style else None,
            outline=para_outline if para_outline is not None else style_outline,
            para_outline=para_outline,
            style_outline=style_outline,
            num_id=num_id, ilvl=int(ilvl) if ilvl is not None else None,
            num_fmt=num_fmt, lvl_text=lvl_text,
            size=size, font=font, run_font=run_font, bold=bold,
            in_table=id(p) in in_tbl,
            form=form,
        ))

    meta = DocxMeta(paras=paras)

    # ---- 目录区块 ----
    # 从「目 录」标题行开始，到下一个带大纲级别/标题样式的段落为止。
    # 这一段的条目在 markdown 里会被清洗规则删掉，若参与对齐会错配
    # （「编制目的」是「一、 编制目的1」的子串）。
    in_toc = False
    toc_end = -1
    for i, p in enumerate(paras):
        t = p.text.strip()
        if not t:
            continue
        if not in_toc:
            if RE_TOC_MARK.match(t) or RE_TOC_MARK.match(t.replace(" ", "")):
                in_toc = True
                p.in_toc = True
                toc_end = i
            continue
        if p.outline is not None or p.md_level is not None:
            in_toc = False
            continue
        p.in_toc = True
        toc_end = i

    # ---- 正文主体起点 ----
    # 封面（公司名 / 文件名 / 版本管理记录）常被作者设成大纲级别 1，
    # 但它们不是章节标题，补进去会把文档骨架搞乱。
    # 主体起点 = 目录之后，或第一个「带大纲级别 / 标题样式 / 编号形态」的段落。
    body_start = len(paras)
    for i, p in enumerate(paras):
        t = p.text.strip()
        if not t:
            continue
        if 0 <= toc_end and i <= toc_end:
            continue
        if p.outline is not None or p.md_level is not None or p.form is not None:
            body_start = i
            break
    meta.body_start = body_start
    for p in paras:
        if p.index < body_start and not p.in_toc:
            p.in_cover = True

    # 编号计数器：numId → {ilvl: 序号}
    # 每段落存一份**当时**的计数快照 —— 还原编号要用该段落自己的序号，
    # 用最终计数会把一组标题全渲染成同一个编号（「4.发起申请 / 4.注销审核」）。
    counters: dict[str, dict[int, int]] = {}
    for p in paras:
        if p.num_id is None:
            continue
        lv = p.ilvl or 0
        c = counters.setdefault(p.num_id, {})
        c[lv] = c.get(lv, 0) + 1
        p.ordinal = c[lv]
        p.counters = dict(c)
    meta.num_counters = counters

    # 编号组 → 形态（取组内多数）
    group_forms: dict[tuple[str, str], Counter] = {}
    for p in paras:
        if p.num_id is None or p.form is None:
            continue
        group_forms.setdefault(p.key, Counter())[p.form] += 1
    for k, c in group_forms.items():
        meta.group_form[k] = c.most_common(1)[0][0]

    # 基准字号：优先取正文样式（Normal）的字号；否则用正文段落的众数
    base_size = None
    for st in (styles.iter(f"{W}style") if styles is not None else ()):
        sid = st.get(f"{W}styleId")
        nm = st.find(f"{W}name")
        if sid is None or nm is None:
            continue
        if (nm.get(f"{W}val") or "").strip().lower() in ("normal", "正文"):
            base_size = style_fmt.get(sid, (None, None))[1]
            break
    if base_size is None:
        body_sizes: dict[float, int] = {}
        all_sizes: dict[float, int] = {}
        for p in paras:
            if (p.in_table or p.outline is not None or p.num_id is not None
                    or p.in_toc or p.in_cover):
                continue
            if not p.text.strip() or p.size is None:
                continue
            all_sizes[p.size] = all_sizes.get(p.size, 0) + 1
            if len(p.text.strip()) > 20:
                body_sizes[p.size] = body_sizes.get(p.size, 0) + 1
        pool = body_sizes or all_sizes
        base_size = max(pool.items(), key=lambda kv: kv[1])[0] if pool else None
    meta.base_size = base_size

    # 标题字体：取「标题样式」定义的字体（heading 1 → 黑体、heading 2 → 楷体）。
    # 不收集所有带大纲级别段落的字体 —— 那些段落的正文会显式设字体，一并收进来就废了。
    head_fonts: set[str] = set()
    for sid in md_levels:
        f, _ = style_fmt.get(sid, (None, None))
        if f:
            head_fonts.add(f)
    for p in paras:
        if p.md_level is not None and p.run_font:
            head_fonts.add(p.run_font)
    meta.head_fonts = frozenset(head_fonts) if head_fonts else STATIC_HEAD_FONTS

    # 编号组判定：组内全部成员都是「短标题」→ 整组是标题组
    groups: dict[tuple[str, str], list[ParaMeta]] = {}
    for p in paras:
        if p.num_id is None or p.in_table or p.in_toc or p.in_cover:
            continue
        groups.setdefault(p.key, []).append(p)
    heading_groups: set[tuple[str, str]] = set()
    for k, members in groups.items():
        usable = [m for m in members if m.text.strip()]
        if len(usable) < 2:
            # 单条组：只有加粗时才当标题
            if len(usable) == 1 and usable[0].bold and _short_title(usable[0].text):
                heading_groups.add(k)
            continue
        if all(_short_title(m.text) for m in usable):
            heading_groups.add(k)
    meta.heading_groups = frozenset(heading_groups)

    return meta


def _short_title(text: str, limit: int = 24) -> bool:
    """短标题形态：不太长、不以句读结尾。"""
    t = text.strip()
    if not t or len(t) > limit:
        return False
    return not (SENT_END & set(t[-1:]))
