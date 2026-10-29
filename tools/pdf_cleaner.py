# -*- coding: utf-8 -*-
"""PDF 专用清洗 —— 去除 MarkItDown 转换 PDF 时引入的页眉/页脚/页码污染。

背景：PDF 没有「页眉页脚」的结构概念，MarkItDown 按文本流顺序抽取，
结果每页重复的页眉（公司名）、页脚（页码）全部混进正文。

典型污染形态：
  - 页眉「某某科技有限公司」每页重复出现
  - 目录页码 6/7/8/9/10 + 罗马页码 I 变成孤立数字行
  - 首页封面区被拆成 `| （内部使用 | 注意保密） |` 这种断裂表格

本模块采用「重复即页眉」的启发式：出现在 >=2 次且本身很短的单行，
且不是章节标题，判定为页眉/页脚并删除。

用法：
    from tools.pdf_cleaner import clean_pdf_markdown
    text = clean_pdf_markdown(raw_text)
"""

import re
from collections import Counter

# 孤立页码行：单独一行的纯数字 / 罗马数字
RE_LONE_PAGE_NO = re.compile(r"^\s*(?:\d{1,4}|[ivxlcdmIVXLCDM]{1,7})\s*$")
# Markdown 表格分隔行（含 --- 或 ----）
RE_TABLE_SEP = re.compile(r"^\|[\s:\-|]+\|$")
# 通用页眉页脚特征词（不含任何公司名——公司名通过调用方参数传入）
#
# 必须是**元组**：写成括号拼接的单个字符串时，`re.escape` 会把其中的 `|`
# 转义成字面量，整个特征词集合退化成一条永远匹配不到的文本
# （实测「编制：中科美络科技股份有限公司」因此清不掉）。
_DEFAULT_HINTS = (
    "股份有限公司",
    "内部使用",
    "注意保密",
    "文件密级",
    "文件编号",
    "文件标识",
    "绝密",
    "机密",
    "秘密",
)
# 章节标题（不能当页眉删）
RE_CHAPTER = re.compile(
    r"^\s*(?:#{1,6}\s|"
    r"(?:第[一二三四五六七八九十]+[章节条]|[一二三四五六七八九十]+\s*[、.．])|"
    r"[（(][一二三四五六七八九十\d]+[)）]|"
    r"附\s*件|"
    r"\d{1,2}\s*[、.．])"
)

# 页码与最近页眉的最大行距。实测这份 PDF 恒为 2；
# 封面年份距页眉 5 行、正文孤立数字距页眉 43 行，取 3 留余量。
_PAGE_NO_HEADER_DISTANCE = 3

# 结构化行前缀：标题 / 图片 / 表格 / 链接 / 引用 —— 不参与页眉页脚判定
_PROTECT_PREFIX = ("#", "!", "|", "<", "[", ">")

# 页眉/页脚行的特征词去掉后，允许残留的「修饰语」长度上限。
# 「编制：中科美络科技股份有限公司」去掉关键词后剩「编制：」→ 剥掉标点剩「编制」（2 字）→ 判定为页脚；
# 「中科美络科技股份有限公司成立于2010年」去掉关键词后剩「成立于2010年」→ 含数字 → 判定为正文，保留。
_MAX_HEADER_RESIDUE = 6
# 判定为正文而非页眉的残留特征：数字（年份、页码、编号）
_RESIDUE_DIGIT = re.compile(r"\d")


def _build_header_pattern(extra_keywords: str = "") -> re.Pattern:
    """合并用户自定义关键词与通用特征词，构造页眉判定正则。

    extra_keywords：用户输入的页眉/页脚文字（如公司名），
    支持逗号/分号/空白分隔多个词。
    """
    words = [w.strip() for w in re.split(r"[,，;；|\s]+", extra_keywords or "") if w.strip()]
    hints = "|".join(re.escape(w) for w in words + list(_DEFAULT_HINTS))
    return re.compile(rf"({hints})")


def _user_keyword_pattern(extra_keywords: str = "") -> re.Pattern | None:
    """只包含用户自定义关键词的正则（不含通用特征词）。"""
    words = [w.strip() for w in re.split(r"[,，;；|\s]+", extra_keywords or "") if w.strip()]
    if not words:
        return None
    return re.compile("|".join(re.escape(w) for w in words))


def _strip_edge_punct(s: str) -> str:
    """剥掉字符串两端的标点与空白。"""
    return s.strip(" \t:：,，。.、;；-—_()（）[]【】<>《》「」\"'“”‘’")


def _is_header_footer_line(s: str, user_pattern: re.Pattern) -> bool:
    """该行是否是「关键词 + 少量修饰语」的页眉/页脚行。

    只看关键词命中是不够的 —— 正文里出现公司名很常见
    （「中科美络科技股份有限公司成立于2010年」）。真正的页眉页脚是
    **整行就是关键词**（可带「编制：」这类前缀），所以要求：
    去掉关键词后，剩余部分剥掉标点后不含数字、且长度很短。
    """
    if not s or len(s) > 80:
        return False
    if s.startswith(_PROTECT_PREFIX):
        return False
    if RE_CHAPTER.match(s):
        return False
    if not user_pattern.search(s):
        return False
    residue = _strip_edge_punct(user_pattern.sub("", s))
    if not residue:
        return True                     # 整行就是关键词
    if len(residue) > _MAX_HEADER_RESIDUE:
        return False
    if _RESIDUE_DIGIT.search(residue):
        return False                    # 残留含数字 → 是正文句子
    return True


def clean_upstream_headers(
    text: str, *, header_keywords: str = ""
) -> tuple[str, dict]:
    """清理**上游解析文本**（text 入口）里的页眉页脚与页码。

    与 ``clean_pdf_markdown`` 的区别（两者判据不同，不能混用）：

    * 上游（MinerU）已经清掉了大部分页眉页脚，**残留的是零星几行** ——
      同一行只出现 1 次。而 ``clean_pdf_markdown`` 的「重复 N 次才算页眉」
      启发式在这里正好失效（实测残留行只出现 1 次 → 清不掉）。
      所以对**用户显式指定的关键词**，改成「命中即删」，但要求整行就是
      关键词（可带少量修饰语），避免误删正文里提到公司名的句子。
    * text 入口的表格已经是 Markdown，**不能删表格分隔行**
      （`clean_pdf_markdown` 会删，那是给 files 入口的 HTML 表格用的）。
    """
    stats = {"headers_removed": 0, "page_nums_removed": 0}
    user_pattern = _user_keyword_pattern(header_keywords)
    if user_pattern is None:
        return text, stats

    lines = text.split("\n")
    removed_idx: list[int] = []
    out: list[tuple[int, str]] = []
    for i, line in enumerate(lines):
        s = line.strip()
        if _is_header_footer_line(s, user_pattern):
            stats["headers_removed"] += 1
            removed_idx.append(i)
            continue
        out.append((i, line))

    # 页码：紧邻被删页眉/页脚（≤3 行）的孤立数字行一并删除。
    # 上游抽取顺序不稳定 —— 页码可能在页脚之前或之后，所以按行距判定。
    kept: list[str] = []
    for i, line in out:
        s = line.strip()
        if RE_LONE_PAGE_NO.match(s) and removed_idx:
            nearest = min(abs(i - h) for h in removed_idx)
            if nearest <= _PAGE_NO_HEADER_DISTANCE:
                stats["page_nums_removed"] += 1
                continue
        kept.append(line)

    result = re.sub(r"\n{3,}", "\n\n", "\n".join(kept))
    return result.strip() + "\n", stats


def _is_candidate_header(
    line: str, freq: dict[str, int], threshold: int, header_hint: re.Pattern
) -> bool:
    """判断某行是否为重复出现的页眉/页脚。"""
    s = line.strip()
    if not s or len(s) > 60:
        return False
    if RE_TABLE_SEP.match(s):
        return False
    if RE_CHAPTER.match(s):
        return False
    # 规则一：含页眉特征词，且重复出现 >= threshold 次
    if header_hint.search(s) and freq.get(s, 0) >= threshold:
        return True
    # 规则二：纯数字行且是连续页码段（6/7/8/9/10 这类）
    if RE_LONE_PAGE_NO.match(s) and freq.get(s, 0) >= 2:
        return True
    return False


def clean_pdf_markdown(
    text: str, *, min_repeat: int = 2, header_keywords: str = ""
) -> tuple[str, dict]:
    """清洗 PDF 转换结果，返回 (清洗后文本, 统计)。

    header_keywords：用户自定义的页眉/页脚特征词（如公司名），
    多个词用逗号分隔。留空则只按通用特征词（密级/编号/保密标识等）判定。
    """
    stats = {"headers_removed": 0, "page_nums_removed": 0, "sep_rows_removed": 0}
    header_hint = _build_header_pattern(header_keywords)

    lines = text.split("\n")
    # 统计每个非空短行出现次数
    freq = Counter(l.strip() for l in lines if l.strip() and len(l.strip()) <= 60)

    # 收集连续页码段（升序连续数字），整体删除
    page_no_idx: set[int] = set()
    i = 0
    while i < len(lines):
        if RE_LONE_PAGE_NO.match(lines[i]):
            start = i
            seq = []
            while i < len(lines) and RE_LONE_PAGE_NO.match(lines[i]):
                seq.append(lines[i].strip())
                i += 1
            # 连续段长度 >= 2 视为页码区（单个数字可能是正文编号）
            if len(seq) >= 2:
                page_no_idx.update(range(start, i))
            continue
        i += 1

    out: list[str] = []
    # 记录「本页页眉/页码」状态。PDF 抽取顺序不稳定：页眉可能在页码之前或之后，
    # 所以页码行只要与已识别页眉相邻（中间只隔空行/表格分隔行）就一并删除。
    removed_header_idx: list[int] = []

    for idx, line in enumerate(lines):
        s = line.strip()
        if idx in page_no_idx:
            stats["page_nums_removed"] += 1
            continue
        if RE_TABLE_SEP.match(s):
            stats["sep_rows_removed"] += 1
            continue
        if _is_candidate_header(line, freq, min_repeat, header_hint):
            stats["headers_removed"] += 1
            removed_header_idx.append(idx)
            continue
        out.append((idx, line))

    # 清理页脚页码。
    # 实测这份 PDF：页码距最近页眉恒为 2 行（页码在视觉页脚，文本流里排在下一页页眉前），
    # 而封面年份 '2025' 距页眉 5 行、正文里的孤立数字距页眉 43 行。
    # 因此用「距离已识别页眉 <= 2 行」作判据，既能删净页码，又不会误删正文数字。
    header_set = set(removed_header_idx)
    kept: list[str] = []
    for idx, line in out:
        s = line.strip()
        if RE_LONE_PAGE_NO.match(s) and header_set:
            nearest = min(abs(idx - h) for h in header_set)
            if nearest <= _PAGE_NO_HEADER_DISTANCE:
                stats["page_nums_removed"] += 1
                continue
        kept.append(line)

    result = re.sub(r"\n{3,}", "\n\n", "\n".join(kept))
    return result.strip() + "\n", stats
