from __future__ import annotations

from PySide6.QtCore import Qt, Signal
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
from excel_helper.ui.widgets import EmptyState


class FieldSelector(QWidget):
    """通用字段选择器组件。

    功能对应产品设计里的 `FieldSelector`：
    - 展示字段名、字段类型、覆盖率、来源文件。
    - 支持全选、取消全选、反选、搜索字段。
    - 向上游页面提供 `selected_fields()`，作为合并/清洗/对比的字段参数。

    上游：主窗口解析文件后调用 `set_fields()`。
    下游：主窗口执行任务前调用 `selected_fields()`。

    界面改造点：
    - 未加载字段时不再留一片空白表格，改为显示空状态引导（`EmptyState`）。
    - 公共字段增加"公共"徽章，部分文件字段用中性徽章，覆盖率一眼可辨。
    - 批量选择按钮统一为 Ghost 按钮，避免和页面主按钮抢视觉权重。
    """

    fieldsChanged = Signal(list)

    def __init__(self, title: str = "选择需要处理的字段") -> None:
        super().__init__()

        # fields 保存后端解析出的字段覆盖率；checkboxes 保存 UI 勾选状态。
        self.fields: list[FieldCoverage] = []
        self.checkboxes: dict[str, QCheckBox] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setObjectName("SectionLabel")
        header.addWidget(title_label)
        header.addStretch(1)

        # 批量选择按钮只影响当前搜索过滤后可见的字段，便于用户局部选择。
        select_all = QPushButton("全选")
        clear_all = QPushButton("取消全选")
        invert = QPushButton("反选")
        for button in (select_all, clear_all, invert):
            button.setObjectName("BtnGhost")
            button.setCursor(Qt.CursorShape.PointingHandCursor)
        select_all.clicked.connect(lambda: self._set_visible_checked(True))
        clear_all.clicked.connect(lambda: self._set_visible_checked(False))
        invert.clicked.connect(self._invert_visible)
        header.addWidget(select_all)
        header.addWidget(clear_all)
        header.addWidget(invert)
        layout.addLayout(header)

        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索字段")
        self.search.setClearButtonEnabled(True)
        # 搜索框输入变化时隐藏不匹配字段，不改变已有勾选状态。
        self.search.textChanged.connect(self._apply_filter)
        layout.addWidget(self.search)

        self.table = QTableWidget(0, 5)
        self.table.setObjectName("FieldTable")
        self.table.setHorizontalHeaderLabels(["选择", "字段", "类型", "覆盖率", "存在文件"])
        # 前四列按内容自适应，来源文件列拉伸占满剩余空间。
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(False)
        layout.addWidget(self.table, 1)

        self.summary = QLabel("已选择：0 / 0")
        self.summary.setObjectName("Muted")
        layout.addWidget(self.summary)

        # 空状态：未加载字段时替代空表格显示。
        self.empty = EmptyState("还没有可选择的字段", "选择文件后会自动解析表头、字段类型与覆盖率")
        layout.addWidget(self.empty, 1)

        self._set_empty_visible(True)

    def set_fields(self, fields: list[FieldCoverage]) -> None:
        """刷新字段列表。

        参数 fields 来自 `build_field_coverage()`，其中 `selected` 决定初始勾选状态：
        公共字段默认选中，部分文件字段默认不选。

        空列表表示"文件未选择或解析失败"，此时切换到空状态，避免用户看到空表格。
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
            # 覆盖率用徽章呈现：全部文件都有 = 绿色公共字段，部分文件有 = 橙色提示。
            badge = QLabel(f"{field.files}/{field.total_files}")
            badge.setObjectName("BadgeSuccess" if field.is_common else "BadgePartial")
            badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setCellWidget(row, 3, badge)
            self.table.setItem(row, 4, QTableWidgetItem("、".join(field.file_names)))
            if field.is_common:
                # 给公共字段打标，后续可以扩展排序或过滤逻辑。
                for column in (1, 2, 4):
                    item = self.table.item(row, column)
                    if item is not None:
                        item.setData(Qt.ItemDataRole.UserRole, "common")
        self._apply_filter()
        self._sync_summary()
        self._set_empty_visible(not fields)
        self.fieldsChanged.emit(self.selected_fields())

    def _set_empty_visible(self, empty: bool) -> None:
        """切换空状态与字段表的显示。

        改造前未选文件时表格是空白的，用户不知道要做什么；现在统一显示引导文案。
        """

        self.empty.setVisible(empty)
        self.table.setVisible(not empty)
        self.search.setVisible(not empty)
        self.summary.setVisible(not empty)

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
        """同步底部统计文案，并通知上游字段勾选发生变化。

        改造前只显示"已选择 x / y"；现在补充公共字段数量，
        让用户在执行前就能确认"所有文件都有的字段"有几个。
        """

        selected = len(self.selected_fields())
        common = sum(1 for field in self.fields if field.is_common)
        self.summary.setText(f"已选择 {selected} / {len(self.fields)} 个字段 · 其中公共字段 {common} 个")
        self.fieldsChanged.emit(self.selected_fields())
