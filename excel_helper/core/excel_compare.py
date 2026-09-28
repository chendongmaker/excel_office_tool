from __future__ import annotations

from pathlib import Path

import pandas as pd

from excel_helper.core.excel_reader import read_table


def compare_by_key(
    left_path: str | Path,
    right_path: str | Path,
    key: str,
    sheet_name: str | int = 0,
    compare_columns: list[str] | None = None,
) -> dict[str, pd.DataFrame]:
    """按主键对比两个 Excel/CSV 文件。

    上游：
    - UI 对比页选择 Excel A/B。
    - 字段选择器返回 `compare_columns`，主键来自“匹配主键”下拉框。

    下游：
    - 返回三张 DataFrame：新增、删除、修改。
    - `ProcessService.compare_files()` 把它们写入同一个 Excel 的三个工作表。
    """

    left = read_table(left_path, sheet_name=sheet_name)
    right = read_table(right_path, sheet_name=sheet_name)
    if key not in left.columns:
        raise ValueError(f"左侧文件缺少对比字段: {key}")
    if key not in right.columns:
        raise ValueError(f"右侧文件缺少对比字段: {key}")

    left_indexed = left.set_index(key, drop=False)
    right_indexed = right.set_index(key, drop=False)

    # 主键集合用于判断新增/删除；交集用于逐字段比较修改。
    left_keys = set(left_indexed.index)
    right_keys = set(right_indexed.index)

    added = right_indexed.loc[sorted(right_keys - left_keys)].reset_index(drop=True)
    removed = left_indexed.loc[sorted(left_keys - right_keys)].reset_index(drop=True)
    if compare_columns:
        # 用户只关心勾选字段时，新增/删除表也只导出主键 + 勾选字段。
        output_columns = [column for column in [key, *compare_columns] if column in added.columns or column in removed.columns]
        added = added[[column for column in output_columns if column in added.columns]]
        removed = removed[[column for column in output_columns if column in removed.columns]]
    changed = _changed_rows(left_indexed, right_indexed, key, sorted(left_keys & right_keys), compare_columns)

    return {"added": added, "removed": removed, "changed": changed}


def _changed_rows(
    left: pd.DataFrame,
    right: pd.DataFrame,
    key: str,
    shared_keys: list[object],
    compare_columns: list[str] | None = None,
) -> pd.DataFrame:
    """生成“修改”明细表。

    每个变化字段输出一行，结构为：
    主键、字段、原值、新值、变化、变化率。
    """

    common_columns = [column for column in left.columns if column in right.columns and column != key]
    if compare_columns:
        # compare_columns 来自字段选择器，只比较用户勾选的业务字段。
        common_columns = [column for column in common_columns if column in compare_columns]
    rows: list[dict[str, object]] = []

    for item_key in shared_keys:
        left_row = left.loc[item_key]
        right_row = right.loc[item_key]
        if isinstance(left_row, pd.DataFrame) or isinstance(right_row, pd.DataFrame):
            # 主键重复时 loc 会返回多行；当前 MVP 先跳过，后续可扩展重复主键报告。
            continue
        for column in common_columns:
            old_value = left_row[column]
            new_value = right_row[column]
            if pd.isna(old_value) and pd.isna(new_value):
                continue
            if old_value != new_value:
                change: dict[str, object] = {
                    key: item_key,
                    "字段": column,
                    "原值": old_value,
                    "新值": new_value,
                }
                if _is_number(old_value) and _is_number(new_value):
                    # 数字字段额外计算绝对变化和变化率，便于销售/库存对比。
                    old_number = float(old_value)
                    new_number = float(new_value)
                    change["变化"] = new_number - old_number
                    change["变化率"] = None if old_number == 0 else (new_number - old_number) / old_number
                rows.append(change)

    return pd.DataFrame(rows)


def _is_number(value: object) -> bool:
    """判断值是否可安全转成数字。"""

    try:
        float(value)
    except (TypeError, ValueError):
        return False
    return not pd.isna(value)
