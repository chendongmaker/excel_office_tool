from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
from pathlib import Path


class TaskStatus(StrEnum):
    """任务执行状态，最终写入历史记录。"""

    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"


@dataclass(slots=True)
class TaskRecord:
    """一次用户操作的历史记录。

    上游：UI 的 `_record_success()` / `_record_failure()` 创建本对象。
    下游：`HistoryStore.add()` 序列化为 JSON，首页和历史任务页再读取展示。
    """

    # 任务名称，例如“纵向合并”“数据清洗”“数据对比”。
    name: str

    # 当前任务状态，成功或失败会展示在历史任务页。
    status: TaskStatus

    # 创建时间，默认记录用户点击执行按钮的时间。
    created_at: datetime = field(default_factory=datetime.now)

    # 输入文件路径列表，用于后续追溯处理来源。
    input_files: list[Path] = field(default_factory=list)

    # 输出文件路径列表，例如导出的 xlsx 或 png。
    outputs: list[Path] = field(default_factory=list)

    # 成功时放导出路径，失败时放错误原因。
    message: str = ""

    def as_dict(self) -> dict[str, object]:
        """转换成可写入 JSON 的普通字典。"""

        return {
            "name": self.name,
            "status": self.status.value,
            "created_at": self.created_at.isoformat(timespec="seconds"),
            "input_files": [str(path) for path in self.input_files],
            "outputs": [str(path) for path in self.outputs],
            "message": self.message,
        }
