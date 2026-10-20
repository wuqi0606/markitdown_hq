# MarkItDown HQ — Dify 插件使用说明

基于微软 [MarkItDown](https://github.com/microsoft/markitdown) 0.1.8 的 Dify 工具插件，
用于 Knowledge Pipeline 的文档处理阶段（Extractor 位置）。

## 解决的核心问题

**标题层级丢失。** 这批企业制度文件的章节标题在 Word 里用的是**普通加粗段落**，
不是「标题1/2/3」样式。MarkItDown 靠样式映射生成 `#`，样式没标就全丢了——
转换后父子分段的 `\n#` / `\n##` 分隔符完全失效。

实测 31 份文件：

| | 补齐前 | 补齐后 |
|---|---|---|
| 零标题（父子分隔符全失效） | 7 份 | **0 份** |
| 只有一级标题（子分隔符失效） | 10 份 | **0 份** |
| 模拟父子分段 · 父分段数 | ~130 | **520** |
| 模拟父子分段 · 子分段数 | ~60 | **695** |

标题补齐已内置（`restore_headings` 参数，默认开启），识别规则：
- `#` → `一、xxx`、`1、xxx`、`附件一、xxx`
- `##` → `（一）xxx`
- `###` → `（1）xxx`

## 与插件市场版 `yevanchen/markitdown` 的区别

| | 市场版 | 本插件 |
|---|---|---|
| MarkItDown 版本 | 0.0.2a1（该版本号 PyPI 不存在，实为 0.0.2） | **0.1.8**（当前稳定版） |
| PDF 支持 | 依赖未声明，PDF 常失败 | `markitdown[docx,pdf]`，PDF 已实测 |
| 标题层级 | 不处理，丢失后无法分段 | **内置补齐** |
| PDF 页眉页脚 | 不处理，混入正文 | **内置清除** |
| 噪声清洗 | 无，输出原样返回 | base64 图片 / 空表格行 / 目录残留 |
| 旧版 .doc 识别 | 报 `BadZipFile`，错误信息不可读 | 按文件头识别，输出明确指引 |
| 多文件 | 拼接为分隔文本 | 逐份处理，失败项单独列出 |

## 安装

### 目录结构要求（踩过坑）

**多媒体文件（图标等）必须放在 `_assets/` 目录下**：

```
dify-markitdown-plugin/
├── _assets/
│   ├── icon.svg          ← 必须在 _assets/ 下
│   └── icon_dark.svg
├── manifest.yaml         ← 这里只写文件名，不带路径
├── main.py
├── requirements.txt
├── provider/
└── tools/
```

manifest 里**只写文件名，不写路径**：

```yaml
icon: icon.svg          # ✅
icon: _assets/icon.svg  # ❌ 会报错
```

放错目录的上传报错：

```
PluginDaemonBadRequestError: file not found: icon.svg
failed to remap plugin icon failed to remap assets failed to save package
```

上传前先跑预检：

```bash
python ../tools/validate_plugin.py ../markitdown_hq.difypkg
```

### 方式一：本地插件调试（推荐先验证）

```bash
# 1. 把官方 CLI 下载到 PATH（macOS/Linux）
curl -L -o dify-plugin https://github.com/langgenius/dify-plugin-daemon/releases/latest/download/dify-plugin-linux-amd64
chmod +x dify-plugin

# 2. 在本目录建虚拟环境并安装 SDK
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install -e https://github.com/langgenius/dify-official-plugins --no-deps  # 或用 dify_plugin SDK

# 3. 配置远程调试
cp .env.example .env
# 编辑 .env 填入 Dify 实例的调试 key

# 4. 打包并在 Dify 中安装
dify-plugin package .
# 在 Dify：插件 → 安装插件 → 通过本地文件
```

### 方式二：使用已打包好的包

`markitdown_hq.difypkg` 已生成。安装：

插件 → 安装插件 → **通过本地文件** → 上传 `.difypkg`

> 注意：`.difypkg` 正式包需要官方 CLI 做签名。若提示签名校验失败，
> 改用方式一的 `dify-plugin package .` 重新打包（CLI 会自动签名）。

## 在 Knowledge Pipeline 中的接法

```
FILE  →  MarkItDown HQ  →  Parent-child Chunker  →  Knowledge Base
```

### MarkItDown HQ 节点参数

| 参数 | 值 | 说明 |
|---|---|---|
| Files | 绑定数据源输出 | 必填 |
| Clean conversion noise | **true** | 知识库场景务必开 |
| **Restore heading levels** | **true** | **知识库场景必须开**，否则父子分隔符失效 |
| **Clean PDF headers/footers** | **true** | PDF 源文件必开，清页眉页脚页码 |
| Add metadata frontmatter | false | 引用显示的是文档名，一般不用 |
| Keep embedded images | **false** | 知识库场景务必关 |

输出开头会有一段 HTML 注释形式的处理报告，只在有异常时出现：

```
<!-- conversion notices
[info] xxx.docx: restored 29 heading(s)
[skip] yyy.docx: legacy Word 97-2003 (.doc) detected ...
-->
```

Dify 会把它当正文索引，但里面是有用信息，别急着删。

### Parent-child Chunker 参数（针对本项目制度文件调优）

**父分段**
- Parent Mode：`Paragraph`（不要用 Full Document，只取前 10000 token 会截断）
- Parent Delimiter：`\n# `
- Parent Max Length：`1200`

**子分段**
- Child Delimiter：`\n\n`
- Child Max Length：`400`

### Knowledge Base 节点

- 索引方式：高质量
- Embedding：`bge-m3` 或 `bge-large-zh`（中文场景）
- 检索方式：混合检索，语义/关键词 0.5:0.5
- Top K：`8`
- Score 阈值：先关闭，确认召回正常后调到 `0.4` 左右
- Rerank：`bge-reranker-v2-m3`

## 输出示例

输入 `项目进度管理办法V1.3.docx` 的版本管理记录表，输出：

```markdown
版本管理记录

|  |  |  |  |  |
| --- | --- | --- | --- | --- |
| **序号** | **日期** | **版本号** | **编写人员** | **编制及变更内容** |
| 1 | 20190315 | V0.1 | 陈振兴 | 初稿 |
| 5 | 20250830 | V1.3 | 刘如霞 | 完善进度偏差处理机制 |
```

正文标题层级保留为 `#` / `##` / `###`，可直接用作父子分段的分隔符。

## 已知限制

- **旧版 `.doc`（Word 97-2003）不支持**。即使后缀是 `.docx`，若文件头是 `D0 CF 11 E0`
  则无法解析，需先用 Word 另存为真正的 `.docx`。插件会在输出开头的
  `<!-- conversion notices -->` 里明确列出。
- 图片被清除后，正文里「如下图所示」这类指代会失去配图。如需保留图片，
  开启 `Keep embedded images`——但知识库场景下 base64 会严重损害向量化质量。
- 目录页残留按「`一、 xxx 3`」和 `[标题](#锚点)` 两种形态清理。若你的文档
  目录格式特殊，可能漏清，检查方式：转换后搜索输出里是否还有带页码的章节标题。
- **PDF 页眉页脚按「重复即页眉」启发式判定**：短行重复出现 ≥2 次且含
  `中科美络`/`股份有限公司`/`文件密级` 等特征词才删。页码则按「距最近已删页眉 ≤3 行」
  判定。实测该 PDF：公司名 8 处 → 1 处（正文那句正确保留），孤立页码 12 → 2
  （残留 2 处分别是封面年份 `2025` 和末页页码，属正常内容）。
  **若你的 PDF 页眉是纯图形或文字高度重复，可能需要调整**
  `tools/pdf_cleaner.py` 里的 `RE_HEADER_HINT` 与 `_PAGE_NO_HEADER_DISTANCE`。
- **PDF 表格会被打散成断裂单元格**，例如
  `| （内部使用 | 注意保密） |     |`。这是 PDF 无表格结构的固有问题，
  本插件无法还原。若某份 PDF 的表格内容重要，建议改用 docx 版或人工校对。

## 本地等价方案

不想装插件的话，本仓库的 `tools/batch_convert.py` + `tools/restore_headings.py`
用同样的逻辑，输出完全一致的 Markdown。

已验证：插件路径与本地路径对 31 份文件的模拟分段结果**均为父 520 / 子 695**。

```bash
python tools/batch_convert.py        # docx -> md
python tools/restore_headings.py     # 补齐标题层级
python tools/smoke_plugin.py         # 插件冒烟测试（7 项）
```