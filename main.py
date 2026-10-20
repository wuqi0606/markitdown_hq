# -*- coding: utf-8 -*-
"""插件入口 —— 标准写法（对照 dify_plugin SDK 0.10.2）。

没有 add_tool 这类注册 API。工具注册链路：
    manifest.yaml (plugins.tools) → tools/markitdown.yaml → extra.python.source
    → tools/markitdown.py (MarkitdownTool 类，SDK 反射发现)
provider 同理：provider/markitdown_hq.yaml → provider/markitdown_hq.py
"""

from dify_plugin import DifyPluginEnv, Plugin

plugin = Plugin(DifyPluginEnv(MAX_REQUEST_TIMEOUT=120))

if __name__ == "__main__":
    plugin.run()