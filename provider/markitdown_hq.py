from dify_plugin import ToolProvider


class MarkitdownHqProvider(ToolProvider):
    """工具提供者。本插件无需凭证，校验直接通过。"""

    def _validate_credentials(self, credentials: dict) -> None:
        pass