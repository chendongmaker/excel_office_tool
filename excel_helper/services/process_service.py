from __future__ import annotations

from pathlib import Path

import pandas as pd

from excel_helper.core.excel_clean import apply_clean_rules
from excel_helper.core.excel_compare import compare_by_key
from excel_helper.core.excel_merge import merge_by_key, merge_by_rows
from excel_helper.core.excel_reader import read_table
from excel_helper.core.excel_writer import write_table
from excel_helper.models.rule import CleanRule


class ProcessService:
    """业务处理服务层。

    这一层是 UI 和 core 之间的“薄胶水”：
    - UI 不直接操作 pandas，也不关心 ExcelWriter。
    - core 只负责 DataFrame 计算，不弹窗、不读 UI 状态。
    - service 把“读取 -> 处理 -> 导出”串成一个完整任务。
    """

    def merge_rows(
        self,
        paths: list[str | Path],
        output_path: str | Path,
        sheet_name: str | int = 0,
        columns: list[str] | None = None,
    ) -> Path:
        """纵向合并并导出 Excel。

        上游变量：
        - paths：合并页文件列表。
        - columns：字段选择器中用户勾选的字段。

        下游：
        - `merge_by_rows()` 返回 DataFrame。
        - `write_table()` 写入“合并结果”工作表。
        """

        frame = merge_by_rows(paths, sheet_name=sheet_name, columns=columns)
        return write_table(frame, output_path, "合并结果")

    def merge_by_key(
        self,
        paths: list[str | Path],
        key: str,
        output_path: str | Path,
        sheet_name: str | int = 0,
        columns: list[str] | None = None,
    ) -> Path:
        """按主键匹配合并并导出 Excel。"""

        frame = merge_by_key(paths, key=key, sheet_name=sheet_name, columns=columns)
        return write_table(frame, output_path, "匹配合并")

    def clean_file(
        self,
        path: str | Path,
        output_path: str | Path,
        rules: list[CleanRule],
        sheet_name: str | int = 0,
        columns: list[str] | None = None,
    ) -> Path:
        """读取单个文件，按字段级规则清洗，再导出 Excel。

        `columns` 先裁剪输入表，`rules` 再逐条作用在这些字段上。
        """

        frame = read_table(path, sheet_name=sheet_name)
        if columns:
            # 只保留用户勾选字段，未勾选字段不会参与清洗或导出。
            frame = frame[[column for column in columns if column in frame.columns]]
        cleaned = apply_clean_rules(frame, rules)
        return write_table(cleaned, output_path, "清洗结果")

    def compare_files(
        self,
        left_path: str | Path,
        right_path: str | Path,
        key: str,
        output_path: str | Path,
        compare_columns: list[str] | None = None,
    ) -> Path:
        """对比两个文件并导出三个工作表。

        上游：
        - key：匹配主键。
        - compare_columns：字段选择器里除主键外的勾选字段。

        下游输出：
        - 新增：右表有、左表没有。
        - 删除：左表有、右表没有。
        - 修改：左右都有同一主键，但字段值不同。
        """

        result = compare_by_key(left_path, right_path, key, compare_columns=compare_columns)
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            # 一个对比任务输出到一个工作簿，便于用户一次性查看三类差异。
            result["added"].to_excel(writer, index=False, sheet_name="新增")
            result["removed"].to_excel(writer, index=False, sheet_name="删除")
            result["changed"].to_excel(writer, index=False, sheet_name="修改")
        return path
