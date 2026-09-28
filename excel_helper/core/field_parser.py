from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

from excel_helper.core.excel_reader import EXCEL_SUFFIXES, read_table
from excel_helper.models.field import ColumnInfo, FieldCoverage, FileMetadata


def parse_file_metadata(path: str | Path, sheet_name: str | int = 0) -> FileMetadata:
    """解析单个 Excel/CSV 文件的表头和字段类型。

    上游：UI 在用户选择文件后调用 `_parse_fields_for_selector()`。
    下游：返回 `FileMetadata`，再由 `build_field_coverage()` 汇总成字段选择器数据。

    参数：
    - path：用户选择的本地文件路径。
    - sheet_name：Excel 工作表名或索引；CSV 会忽略工作表概念。
    """

    file_path = Path(path)

    # 这里只读取 DataFrame 的结构信息，不在 UI 中保留完整数据，避免页面状态过重。
    frame = read_table(file_path, sheet_name=sheet_name)
    display_sheet = _sheet_display_name(file_path, sheet_name)

    # pandas dtype 是后续“字段类型提示”的来源，帮助用户区分金额、日期、文本等列。
    columns = [ColumnInfo(str(column), str(dtype)) for column, dtype in frame.dtypes.items()]
    return FileMetadata(
        file_id=_file_id(file_path),
        file_name=file_path.name,
        path=file_path,
        sheet_name=display_sheet,
        columns=columns,
    )


def parse_files_metadata(paths: list[str | Path], sheet_name: str | int = 0) -> list[FileMetadata]:
    """批量解析多个文件。

    上游：合并页、清洗页、对比页选择文件后调用。
    下游：返回每个文件的字段列表，用于计算公共字段和部分字段。
    """

    return [parse_file_metadata(path, sheet_name=sheet_name) for path in paths]


def build_field_coverage(files: list[FileMetadata]) -> list[FieldCoverage]:
    """统计多个文件中每个字段的覆盖率。

    功能对应需求里的“公共字段 + 仅部分文件存在”：
    - `files == total_files`：公共字段，默认勾选。
    - `files < total_files`：部分字段，默认不勾选。

    下游：`FieldSelector.set_fields()` 直接使用返回列表渲染表格。
    """

    total = len(files)

    # names 记录字段出现在哪些文件中；dtypes 记录同名字段在不同文件里的类型分布。
    names: dict[str, list[str]] = defaultdict(list)
    dtypes: dict[str, Counter[str]] = defaultdict(Counter)

    for file_meta in files:
        seen_in_file: set[str] = set()
        for column in file_meta.columns:
            # 同一个文件里如果有重复列名，只计一次覆盖率，避免 1 个文件被算成多份。
            if column.name in seen_in_file:
                continue
            seen_in_file.add(column.name)
            names[column.name].append(file_meta.file_name)
            dtypes[column.name][column.dtype] += 1

    coverage = [
        FieldCoverage(
            name=name,
            # 如果同名字段在多个文件类型不同，取最常见的类型做 UI 提示。
            dtype=dtypes[name].most_common(1)[0][0],
            files=len(file_names),
            total_files=total,
            file_names=file_names,
            selected=len(file_names) == total,
        )
        for name, file_names in names.items()
    ]
    return sorted(coverage, key=lambda item: (not item.is_common, item.name))


def selected_columns_from_coverage(fields: list[FieldCoverage]) -> list[str]:
    """从字段覆盖率数据中取出已勾选字段。

    当前 UI 直接从 `FieldSelector` 读取勾选状态；这个函数保留给后续 API
    或非 UI 流程复用。
    """

    return [field.name for field in fields if field.selected]


def _sheet_display_name(path: Path, sheet_name: str | int) -> str:
    """把工作表索引转换成用户可读名称。"""

    if isinstance(sheet_name, str):
        return sheet_name
    if path.suffix.lower() not in EXCEL_SUFFIXES:
        return "CSV"
    excel = pd.ExcelFile(path)
    try:
        return str(excel.sheet_names[sheet_name])
    except IndexError:
        return str(sheet_name)


def _file_id(path: Path) -> str:
    """根据本地路径生成文件 ID。

    现在是桌面版本地文件 ID；如果后续接 FastAPI 上传接口，这里可以替换为服务端 ID。
    """

    return f"file_{abs(hash(path.resolve())):x}"
