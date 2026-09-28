from __future__ import annotations

from pathlib import Path

import pandas as pd

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}
TABLE_SUFFIXES = EXCEL_SUFFIXES | {".csv"}


def list_table_files(folder: str | Path) -> list[Path]:
    """列出一个文件夹下支持处理的表格文件。

    上游：后续如果做“选择文件夹批处理”，可以直接调用本函数。
    下游：返回路径列表，再交给 `read_tables()` 或字段解析服务。
    """

    base = Path(folder)
    if not base.exists():
        raise FileNotFoundError(f"文件夹不存在: {base}")
    return sorted(path for path in base.iterdir() if path.suffix.lower() in TABLE_SUFFIXES)


def read_table(path: str | Path, sheet_name: str | int = 0) -> pd.DataFrame:
    """读取单个 Excel/CSV 文件为 pandas DataFrame。

    这是所有数据处理功能的最底层入口：
    - 字段解析：读取后拿 `columns` 和 `dtypes`。
    - 合并/清洗/对比：读取后交给 pandas 处理。
    - 图表：读取后按用户选择字段聚合绘图。

    参数：
    - path：本地表格文件路径。
    - sheet_name：Excel 工作表名或索引；CSV 会直接读取整份文件。
    """

    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(file_path)
    if suffix in EXCEL_SUFFIXES:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    raise ValueError(f"不支持的文件类型: {file_path.suffix}")


def select_existing_columns(frame: pd.DataFrame, columns: list[str] | None, required: list[str] | None = None) -> pd.DataFrame:
    """按用户选择字段裁剪 DataFrame，同时保留必需字段。

    上游变量：
    - columns：来自 `FieldSelector.selected_fields()`，代表用户勾选的字段。
    - required：来自合并/对比逻辑，例如按主键合并时主键必须存在。

    下游：返回裁剪后的 DataFrame，减少后续处理和导出列数。
    """

    if not columns:
        return frame
    required = required or []

    # dict.fromkeys 用于去重并保持顺序，保证必需字段排在用户字段前面。
    selected = list(dict.fromkeys([*required, *columns]))
    return frame[[column for column in selected if column in frame.columns]]


def read_tables(
    paths: list[str | Path],
    sheet_name: str | int = 0,
    include_source: bool = True,
    columns: list[str] | None = None,
    required_columns: list[str] | None = None,
) -> list[pd.DataFrame]:
    """批量读取多个表格文件。

    上游：`excel_merge` 调用本函数读取多文件。
    下游：返回 DataFrame 列表，供 `pd.concat()` 或 `pd.merge()` 使用。

    参数：
    - include_source：为纵向合并添加“来源文件”，方便追溯每行来自哪个 Excel。
    - columns：用户在字段选择器中勾选的字段。
    - required_columns：任务必须保留的字段，例如匹配主键。
    """

    frames: list[pd.DataFrame] = []
    for path in paths:
        file_path = Path(path)
        frame = read_table(file_path, sheet_name=sheet_name)
        frame = select_existing_columns(frame, columns, required_columns)
        if include_source:
            # 来源文件放在第一列，让合并后的结果更容易审计。
            frame.insert(0, "来源文件", file_path.name)
        frames.append(frame)
    return frames
