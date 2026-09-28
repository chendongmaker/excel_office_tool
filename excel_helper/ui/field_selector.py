from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_helper.models.field import FieldCoverage


class FieldSelector(QWidget):
    """通用字段选择器组件。

    功能对应产品设计里的 `FieldSelector`：
    - 展示字段名、字段类型、覆盖率、来源文件。
    - 支持全选、取消全选、反选、搜索字段。
    - 向上游页面提供 `selected_fields()`，作为合并/清洗/对比的字段参数。

    上游：主窗口解析文件后调用 `set_fields()`。
    下游：主窗口执行任务前调用 `selected_fields()`。
    """

    def __init__(self, title: str = "选择需要处理的字段") -> None:
        super().__init__()

        # fields 保存后端解析出的字段覆盖率；checkboxes 保存 UI 勾选状态。
        self.fields: list[FieldCoverage] = []
        self.checkboxes: dict[str, QCheckBox] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        header = QHBoxLayout()
        header.addWidget(QLabel(title))
        header.addStretch(1)

        # 批量选择按钮只影响当前搜索过滤后可见的字段，便于用户局部选择。
        select_all = QPushButton("全选")
        clear_all = QPushButton("取消全选")
        invert = QPushButton("反选")
        select_all.clicked.connect(lambda: self._set_visible_checked(True))
        clear_all.clicked.connect(lambda: self._set_visible_checked(False))
        invert.clicked.connect(self._invert_visible)
        header.addWidget(select_all)
        header.addWidget(clear_all)
        header.addWidget(invert)
        layout.addLayout(header)

        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索字段")
        # 搜索框输入变化时隐藏不匹配字段，不改变已有勾选状态。
        self.search.textChanged.connect(self._apply_filter)
        layout.addWidget(self.search)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["选择", "字段", "类型", "覆盖率", "存在文件"])
        # 前四列按内容自适应，来源文件列拉伸占满剩余空间。
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table, 1)

        self.summary = QLabel("已选择：0 / 0")
        layout.addWidget(self.summary)

    def set_fields(self, fields: list[FieldCoverage]) -> None:
        """刷新字段列表。

        参数 fields 来自 `build_field_coverage()`，其中 `selected` 决定初始勾选状态：
        公共字段默认选中，部分文件字段默认不选。
        """

        self.fields = fields
        self.checkboxes = {}
        self.table.setRowCount(len(fields))
        for row, field in enumerate(fields):
            # 每行第一列放真正的 QCheckBox，后续读取用户选择时从这里取值。
            checkbox = QCheckBox()
            checkbox.setChecked(field.selected)
            checkbox.stateChanged.connect(self._sync_summary)
            self.checkboxes[field.name] = checkbox
            self.table.setCellWidget(row, 0, checkbox)
            self.table.setItem(row, 1, QTableWidgetItem(field.name))
            self.table.setItem(row, 2, QTableWidgetItem(field.dtype))
            self.table.setItem(row, 3, QTableWidgetItem(f"{field.files}/{field.total_files}"))
            self.table.setItem(row, 4, QTableWidgetItem("、".join(field.file_names)))
            if field.is_common:
                # 给公共字段打标，后续可以扩展颜色样式或排序逻辑。
                for column in range(1, 5):
                    self.table.item(row, column).setData(Qt.ItemDataRole.UserRole, "common")
        self._apply_filter()
        self._sync_summary()

    def selected_fields(self) -> list[str]:
        """返回当前用户勾选的字段名列表。

        上游页面会把这个列表传给：
        - `ProcessService.merge_rows(columns=...)`
        - `ProcessService.clean_file(columns=...)`
        - `ProcessService.compare_files(compare_columns=...)`
        """

        selected: list[str] = []
        for field in self.fields:
            checkbox = self.checkboxes.get(field.name)
            if checkbox and checkbox.isChecked():
                selected.append(field.name)
        return selected

    def has_fields(self) -> bool:
        """当前组件是否已经加载过字段。"""

        return bool(self.fields)

    def _set_visible_checked(self, checked: bool) -> None:
        """设置当前可见字段的勾选状态。"""

        for row, field in enumerate(self.fields):
            if self.table.isRowHidden(row):
                continue
            self.checkboxes[field.name].setChecked(checked)
        self._sync_summary()

    def _invert_visible(self) -> None:
        """反选当前可见字段。"""

        for row, field in enumerate(self.fields):
            if self.table.isRowHidden(row):
                continue
            checkbox = self.checkboxes[field.name]
            checkbox.setChecked(not checkbox.isChecked())
        self._sync_summary()

    def _apply_filter(self) -> None:
        """根据搜索关键字显示/隐藏字段行。"""

        keyword = self.search.text().strip().lower()
        for row, field in enumerate(self.fields):
            self.table.setRowHidden(row, bool(keyword and keyword not in field.name.lower()))

    def _sync_summary(self) -> None:
        """同步底部“已选择 x / y”的统计文案。"""

        selected = len(self.selected_fields())
        self.summary.setText(f"已选择：{selected} / {len(self.fields)}")
