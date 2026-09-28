from __future__ import annotations

from functools import reduce
from pathlib import Path

import pandas as pd

from excel_helper.core.excel_reader import read_tables


def merge_by_rows(
    paths: list[str | Path],
    sheet_name: str | int = 0,
    include_source: bool = True,
    columns: list[str] | None = None,
) -> pd.DataFrame:
    """按行纵向合并多个 Excel/CSV。

    上游：
    - UI 合并页选择文件并解析字段。
    - `ProcessService.merge_rows()` 传入用户勾选的 `columns`。

    下游：
    - 返回合并后的 DataFrame。
    - service 层再调用 `write_table()` 导出为 Excel。
    """

    frames = read_tables(paths, sheet_name=sheet_name, include_source=include_source, columns=columns)
    if not frames:
        raise ValueError("请至少选择一个 Excel 或 CSV 文件")

    # sort=False 保持原始字段顺序；不同文件缺少的列由 pandas 填 NaN。
    return pd.concat(frames, ignore_index=True, sort=False)


def merge_by_key(
    paths: list[str | Path],
    key: str,
    sheet_name: str | int = 0,
    how: str = "outer",
    columns: list[str] | None = None,
) -> pd.DataFrame:
    """按指定主键横向匹配合并多个文件。

    典型场景：
    - 文件 A 有“商品ID、商品名”
    - 文件 B 有“商品ID、销量”
    - 用户选择商品ID作为 key，结果按商品ID合成一张宽表。

    参数：
    - key：匹配字段，来自 UI 的“匹配字段/匹配主键”输入。
    - columns：字段选择器勾选字段，读取时会额外强制保留 key。
    - how：传给 `pandas.merge()` 的连接方式，默认 outer 保留两侧全部数据。
    """

    frames = read_tables(paths, sheet_name=sheet_name, include_source=False, columns=columns, required_columns=[key])
    if len(frames) < 2:
        raise ValueError("按字段匹配合并至少需要两个文件")

    # 在真正 merge 前先给出清晰错误，避免 pandas 报错对办公用户不友好。
    for frame, path in zip(frames, paths, strict=True):
        if key not in frame.columns:
            raise ValueError(f"{Path(path).name} 缺少匹配字段: {key}")

    # reduce 会按顺序把多个 DataFrame 两两 merge，支持 2 个以上文件。
    return reduce(lambda left, right: pd.merge(left, right, on=key, how=how), frames)
