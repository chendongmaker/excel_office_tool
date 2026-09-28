from __future__ import annotations

from pathlib import Path

from excel_helper.chart.chart_factory import save_chart
from excel_helper.core.excel_reader import read_table


class ChartService:
    """图表业务服务层。

    上游：UI 图表页传入文件、图表类型、X/Y 字段、分组字段。
    下游：读取表格后交给 `chart_factory.save_chart()` 生成 PNG。
    """

    def create_chart(
        self,
        input_path: str | Path,
        output_path: str | Path,
        chart_type: str,
        x_column: str,
        y_column: str,
        group_column: str | None = None,
        title: str = "",
    ) -> Path:
        """读取数据源并生成图表文件。"""

        frame = read_table(input_path)
        return save_chart(frame, chart_type, x_column, y_column, output_path, group_column, title)
