from __future__ import annotations

import json
from pathlib import Path

from excel_helper.models.task import TaskRecord


class HistoryStore:
    """本地历史记录存储。

    当前 MVP 使用 JSON 文件，路径默认在用户主目录下：
    `~/.excel_batch_assistant/history.json`。

    上游：主窗口任务成功/失败时调用 `add()`。
    下游：首页最近任务和“历史任务”页面调用 `list()` 展示。
    """

    def __init__(self, path: str | Path | None = None) -> None:
        """初始化历史文件路径，并确保父目录存在。"""

        self.path = Path(path or Path.home() / ".excel_batch_assistant" / "history.json")
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def add(self, record: TaskRecord) -> None:
        """新增一条任务记录。

        最新记录放在最前面，只保留最近 100 条，防止历史文件无限增长。
        """

        items = self.list()
        items.insert(0, record.as_dict())
        self.path.write_text(json.dumps(items[:100], ensure_ascii=False, indent=2), encoding="utf-8")

    def list(self) -> list[dict[str, object]]:
        """读取历史记录。

        如果文件不存在或 JSON 损坏，返回空列表，保证 UI 不会因为历史文件问题崩溃。
        """

        if not self.path.exists():
            return []
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return []
