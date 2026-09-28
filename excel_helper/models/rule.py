from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum


class CleanAction(StrEnum):
    """数据清洗动作枚举。

    UI 根据用户勾选字段和输入项组装 `CleanRule`；
    `core.excel_clean.apply_clean_rules()` 再按这里的动作逐条执行。
    """

    DROP_EMPTY_ROWS = "drop_empty_rows"
    DROP_DUPLICATES = "drop_duplicates"
    RENAME_COLUMNS = "rename_columns"
    DROP_COLUMNS = "drop_columns"
    FILL_EMPTY = "fill_empty"
    STRIP_TEXT = "strip_text"
    NORMALIZE_DATES = "normalize_dates"
    NORMALIZE_NUMBERS = "normalize_numbers"


@dataclass(slots=True)
class CleanRule:
    """一条字段级清洗规则。

    上游：主窗口 `_run_clean()` 根据字段选择器和表单输入生成规则列表。
    下游：`apply_clean_rules()` 顺序执行规则，每条规则都会返回新的 DataFrame。
    """

    # 要执行的清洗动作，必须来自 `CleanAction`。
    action: CleanAction

    # 允许后续做规则编辑器时临时禁用某条规则。
    enabled: bool = True

    # 规则影响的字段列表；为空时部分规则会默认作用于全部字段。
    columns: list[str] = field(default_factory=list)

    # 改列名规则专用：旧列名 -> 新列名。
    rename_map: dict[str, str] = field(default_factory=dict)

    # 空值填充规则专用，当前由 UI 的“空值填充”输入框传入。
    fill_value: str | int | float = ""

    # 日期标准化输出格式，默认转为 2026-08-27 这种形式。
    date_format: str = "%Y-%m-%d"
