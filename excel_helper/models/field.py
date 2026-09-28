from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True)
class ColumnInfo:
    """单个字段的解析结果。

    上游：`field_parser.parse_file_metadata()` 从 pandas DataFrame 的
    `columns` 和 `dtypes` 生成本模型。
    下游：`FieldCoverage` 汇总多个文件里的同名字段后，交给 UI 的
    `FieldSelector` 展示。
    """

    # Excel/CSV 表头中的原始字段名，会直接显示给用户，也会作为后续处理的列名。
    name: str

    # pandas 推断出的字段类型，例如 int64、float64、object、string。
    dtype: str


@dataclass(slots=True)
class FileMetadata:
    """单个文件的元数据。

    它代表“上传/选择文件后，后端解析表头”的返回结构。
    UI 页面只依赖这里的轻量信息，不直接读取完整 DataFrame。
    """

    # 稳定的文件标识，后续如果改成 FastAPI 上传接口，可以映射到服务端临时文件。
    file_id: str

    # 用户可读的文件名，用于文件列表、字段覆盖率来源提示和操作日志。
    file_name: str

    # 本地真实路径；当前桌面版直接用它读取文件。
    path: Path

    # 当前解析的工作表名称；CSV 统一显示为 CSV。
    sheet_name: str

    # 当前文件中所有字段的名称和类型。
    columns: list[ColumnInfo]


@dataclass(slots=True)
class FieldCoverage:
    """多个文件合并后的字段覆盖率。

    上游：`build_field_coverage()` 统计所有 `FileMetadata.columns`。
    下游：`FieldSelector` 根据 `selected` 初始勾选状态渲染表格；
    执行任务时 UI 再调用 `FieldSelector.selected_fields()` 取回用户选择。
    """

    # 字段名，作为最终传给合并、清洗、对比逻辑的列名。
    name: str

    # 多个文件中该字段最常见的 pandas dtype，用于辅助用户判断字段含义。
    dtype: str

    # 有多少个文件包含该字段。
    files: int

    # 本次任务总文件数，用于显示“3/5 文件存在”。
    total_files: int

    # 包含该字段的文件名列表，用于 UI 的“存在文件”列。
    file_names: list[str] = field(default_factory=list)

    # 初始是否勾选。默认公共字段勾选，部分文件字段不勾选，减少误合并风险。
    selected: bool = True

    @property
    def is_common(self) -> bool:
        """字段是否存在于所有文件中。"""

        return self.files == self.total_files
