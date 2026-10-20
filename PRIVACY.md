# PRIVACY.md

## 数据处理说明

**本插件不会将任何数据发送到外部服务。**

所有文档转换均在 Dify 插件运行沙箱内本地完成：

- 文档内容由 Microsoft MarkItDown 库在本地进程内解析
- 不调用任何远程 API、不上传文件、不记录文档内容到日志
- 插件不声明 `tool` 之外的任何模型调用权限（无 LLM、无 Embedding、无 Rerank）

## 依赖

- `markitdown[docx,pdf]==0.1.8` —— Microsoft 开源库，MIT 协议
  - 仅使用其本地文档解析能力（mammoth / pdfminer 等）
  - 已通过 `enable_plugins=False` 关闭其可选插件机制，避免触发任何网络请求

## 权限

- `tool.enabled: true` —— 允许在 Workflow / Knowledge Pipeline 中被调用
- 无其他权限申请（不申请模型推理、不申请持久化存储）