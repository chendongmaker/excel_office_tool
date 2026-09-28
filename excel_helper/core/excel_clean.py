from __future__ import annotations

import re

import pandas as pd

from excel_helper.models.rule import CleanAction, CleanRule


def apply_clean_rules(frame: pd.DataFrame, rules: list[CleanRule]) -> pd.DataFrame:
    """按顺序执行清洗规则。

    上游：`ProcessService.clean_file()` 读取 Excel 后传入 DataFrame 和规则列表。
    下游：返回清洗后的 DataFrame，再导出到 Excel。

    规则顺序很重要，例如先去空格再改列名，可以让用户输入的字段更统一。
    """

    cleaned = frame.copy()
    for rule in rules:
        if not rule.enabled:
            continue
        cleaned = _apply_rule(cleaned, rule)
    return cleaned


def _apply_rule(frame: pd.DataFrame, rule: CleanRule) -> pd.DataFrame:
    """执行单条清洗规则，并返回新的 DataFrame。"""

    if rule.action == CleanAction.DROP_EMPTY_ROWS:
        # 删除整行全为空的数据，避免合并/统计时出现无意义空行。
        return frame.dropna(how="all")
    if rule.action == CleanAction.DROP_DUPLICATES:
        subset = _target_columns(frame, rule.columns) or None
        # subset 来自字段选择器；为空时 pandas 会按整行判断重复。
        return frame.drop_duplicates(subset=subset)
    if rule.action == CleanAction.RENAME_COLUMNS:
        # rename_map 来自 UI 的“旧列名=新列名”输入框。
        return frame.rename(columns=rule.rename_map)
    if rule.action == CleanAction.DROP_COLUMNS:
        # 只删除实际存在的列，用户输入不存在列时不让任务失败。
        return frame.drop(columns=[column for column in rule.columns if column in frame.columns])
    if rule.action == CleanAction.FILL_EMPTY:
        target_columns = _target_columns(frame, rule.columns)
        if rule.columns:
            frame = frame.copy()
            frame[target_columns] = frame[target_columns].fillna(rule.fill_value)
            return frame
        return frame.fillna(rule.fill_value)
    if rule.action == CleanAction.STRIP_TEXT:
        return _strip_text(frame, rule.columns)
    if rule.action == CleanAction.NORMALIZE_DATES:
        return _normalize_dates(frame, rule.columns, rule.date_format)
    if rule.action == CleanAction.NORMALIZE_NUMBERS:
        return _normalize_numbers(frame, rule.columns)
    return frame


def _target_columns(frame: pd.DataFrame, columns: list[str]) -> list[str]:
    """把用户选择的字段名映射到当前 DataFrame 中实际存在的列。

    这里额外兼容列名前后空格：用户看到/输入“ 商品 ”时，前面清洗步骤
    可能已经把列名变成“商品”，因此需要二次匹配。
    """

    target: list[str] = []
    for column in columns or list(frame.columns):
        if column in frame.columns:
            target.append(column)
            continue
        stripped = column.strip() if isinstance(column, str) else column
        if stripped in frame.columns:
            target.append(stripped)
    return list(dict.fromkeys(target))


def _strip_text(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """去除文本字段和字段名的前后空格。"""

    result = frame.copy()

    # 字段名也做 trim，解决 Excel 表头里常见的隐藏空格问题。
    result.columns = [column.strip() if isinstance(column, str) else column for column in result.columns]
    for column in _target_columns(frame, columns):
        if column not in result.columns:
            stripped_column = column.strip() if isinstance(column, str) else column
            if stripped_column not in result.columns:
                continue
            column = stripped_column
        if pd.api.types.is_string_dtype(result[column]) or result[column].dtype == "object":
            # astype(str) 让数字/空值混杂的文本列也能统一调用 str.strip()。
            result[column] = result[column].astype(str).str.strip()
    return result


def _normalize_dates(frame: pd.DataFrame, columns: list[str], date_format: str) -> pd.DataFrame:
    """把用户选择字段中可识别的日期统一格式。"""

    result = frame.copy()
    for column in _target_columns(frame, columns):
        # errors="coerce" 会把无法识别的值转为 NaT，后面再恢复原值。
        values = pd.to_datetime(result[column], errors="coerce")
        result[column] = values.dt.strftime(date_format)
        # 无法识别为日期的单元格保持原样，避免误删“暂无”等业务文本。
        result.loc[values.isna(), column] = frame.loc[values.isna(), column]
    return result


def _normalize_numbers(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """把金额/数字文本尽量转换为数字。

    示例：`￥1,200元` -> `1200`。转换失败的值会保留原文，方便用户排查。
    """

    result = frame.copy()
    for column in _target_columns(frame, columns):
        if pd.api.types.is_numeric_dtype(result[column]):
            continue
        # 去掉常见金额符号、千分位逗号、空白和“元”字。
        cleaned = (
            result[column]
            .astype(str)
            .map(lambda value: re.sub(r"[,\s￥¥元]", "", value))
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        )
        numeric = pd.to_numeric(cleaned, errors="coerce")
        # 能转数字的用数字，不能转的保留原值。
        result[column] = numeric.where(numeric.notna(), result[column])
    return result
