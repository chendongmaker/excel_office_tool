from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_helper.core.field_parser import build_field_coverage, parse_files_metadata
from excel_helper.models.rule import CleanAction, CleanRule
from excel_helper.models.task import TaskRecord, TaskStatus
from excel_helper.services.chart_service import ChartService
from excel_helper.services.process_service import ProcessService
from excel_helper.ui.field_selector import FieldSelector
from excel_helper.storage.history import HistoryStore


class MainWindow(QMainWindow):
    """Excel批处理助手主窗口。

    主窗口负责三件事：
    1. 创建页面和控件，收集用户输入。
    2. 在用户选择文件后调用 `field_parser` 解析字段，并刷新 `FieldSelector`。
    3. 在用户点击执行按钮后调用 service 层完成处理、导出和历史记录。

    分层关系：
    UI(MainWindow/FieldSelector) -> Service(ProcessService/ChartService) -> Core(pandas/openpyxl/matplotlib)
    """

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Excel批处理助手")
        self.resize(1180, 760)

        # service 层封装“读取 -> 处理 -> 导出”，UI 不直接写 pandas 逻辑。
        self.process_service = ProcessService()
        self.chart_service = ChartService()

        # history 负责本地 JSON 历史记录，供首页和历史任务页展示。
        self.history = HistoryStore()

        # 左侧导航 + 右侧堆叠页面，是整个桌面端的主布局。
        self.stack = QStackedWidget()
        self.nav = QListWidget()
        self.nav.addItems(["首页", "批量合并", "数据清洗", "数据对比", "图表生成", "历史任务"])
        self.nav.setFixedWidth(180)
        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)

        self.stack.addWidget(self._home_page())
        self.stack.addWidget(self._merge_page())
        self.stack.addWidget(self._clean_page())
        self.stack.addWidget(self._compare_page())
        self.stack.addWidget(self._chart_page())
        self.stack.addWidget(self._history_page())

        root = QWidget()
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.nav)
        layout.addWidget(self.stack, 1)
        self.setCentralWidget(root)

        refresh_action = QAction("刷新历史", self)
        refresh_action.triggered.connect(self._refresh_history)
        self.menuBar().addAction(refresh_action)
        self.nav.setCurrentRow(0)
        self._apply_style()

    def _home_page(self) -> QWidget:
        """创建首页。

        首页是功能入口和最近任务概览，不执行数据处理。
        """

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(32, 28, 32, 28)
        title = QLabel("Excel批处理助手")
        title.setObjectName("Title")
        subtitle = QLabel("拖入文件，批量处理，自动对比，生成图表")
        subtitle.setObjectName("Subtitle")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        cards = QFrame()
        grid = QHBoxLayout(cards)
        grid.setSpacing(14)
        for text, index in [
            ("批量合并Excel", 1),
            ("数据清洗", 2),
            ("数据对比", 3),
            ("生成图表", 4),
        ]:
            button = QPushButton(text)
            button.setMinimumHeight(120)
            button.setObjectName("CardButton")
            # 点击功能卡片时切换到对应页面，index 与 nav/stack 中的页面顺序一致。
            button.clicked.connect(lambda checked=False, row=index: self.nav.setCurrentRow(row))
            grid.addWidget(button)
        layout.addWidget(cards)
        layout.addWidget(QLabel("最近任务"))
        self.home_history = QListWidget()
        layout.addWidget(self.home_history, 1)
        self._refresh_home_history()
        return page

    def _merge_page(self) -> QWidget:
        """创建“批量合并”页面。

        页面数据流：
        选择文件 -> 解析表头 -> FieldSelector 选字段 -> 选择合并方式 -> ProcessService 导出。
        """

        page = QWidget()
        layout = self._page_layout(page, "批量合并Excel")
        self.merge_files = QListWidget()

        # 合并页字段选择器展示所有文件的公共字段和部分字段覆盖率。
        self.merge_field_selector = FieldSelector("② 选择字段")
        self.merge_log = QListWidget()
        self.merge_mode_group = QButtonGroup(page)
        self.merge_rows_radio = QRadioButton("按行纵向合并")
        self.merge_key_radio = QRadioButton("按指定字段匹配")
        self.merge_rows_radio.setChecked(True)
        self.merge_mode_group.addButton(self.merge_rows_radio)
        self.merge_mode_group.addButton(self.merge_key_radio)
        self.merge_key_input = QLineEdit()
        self.merge_key_input.setPlaceholderText("例如：商品ID")

        pick = QPushButton("选择Excel或CSV文件")
        # 选择后不直接执行合并，而是先解析字段供用户勾选。
        pick.clicked.connect(self._choose_merge_files)
        run = QPushButton("开始合并")
        run.clicked.connect(self._run_merge)

        layout.addWidget(pick)
        layout.addWidget(self.merge_files)
        layout.addWidget(self.merge_field_selector, 1)
        layout.addWidget(self.merge_rows_radio)
        layout.addWidget(self.merge_key_radio)
        form = QFormLayout()
        form.addRow("匹配字段", self.merge_key_input)
        layout.addLayout(form)
        layout.addWidget(QLabel("操作日志"))
        layout.addWidget(self.merge_log)
        layout.addWidget(run)
        return page

    def _clean_page(self) -> QWidget:
        """创建“数据清洗”页面。

        页面数据流：
        选择单个文件 -> 解析表头 -> 选择需要清洗的字段 -> 组装 CleanRule -> ProcessService 导出。
        """

        page = QWidget()
        layout = self._page_layout(page, "批量数据清洗")
        self.clean_file_input = QLineEdit()
        self.clean_file_input.setReadOnly(True)

        # 清洗页字段选择器决定哪些列参与清洗和导出。
        self.clean_field_selector = FieldSelector("选择需要清洗的字段")
        self.clean_log = QListWidget()
        self.rename_input = QLineEdit()
        self.rename_input.setPlaceholderText("旧列名=新列名，多条用逗号分隔")
        self.drop_columns_input = QLineEdit()
        self.drop_columns_input.setPlaceholderText("成本价,内部备注")
        self.fill_value_input = QLineEdit("0")

        pick = QPushButton("选择待清洗文件")
        pick.clicked.connect(self._choose_clean_file)
        form = QFormLayout()
        form.addRow("文件", self.clean_file_input)
        form.addRow("修改列名", self.rename_input)
        form.addRow("删除列", self.drop_columns_input)
        form.addRow("空值填充", self.fill_value_input)
        run = QPushButton("开始清洗")
        run.clicked.connect(self._run_clean)
        layout.addWidget(pick)
        layout.addLayout(form)
        layout.addWidget(self.clean_field_selector, 1)
        layout.addWidget(QLabel("默认会执行：删除空行、删除重复行、文本去空格、日期标准化、数字标准化。"))
        layout.addWidget(QLabel("操作日志"))
        layout.addWidget(self.clean_log)
        layout.addWidget(run)
        return page

    def _compare_page(self) -> QWidget:
        """创建“数据对比”页面。

        页面数据流：
        选择 Excel A/B -> 解析公共字段 -> 选择匹配主键 -> 勾选对比字段 -> 导出差异报告。
        """

        page = QWidget()
        layout = self._page_layout(page, "Excel数据对比")
        self.compare_left_input = QLineEdit()
        self.compare_right_input = QLineEdit()
        self.compare_key_input = QComboBox()
        self.compare_key_input.setEditable(True)

        # 对比页字段选择器表示“要比较哪些字段”，主键通过单独下拉框选择。
        self.compare_field_selector = FieldSelector("选择需要对比的字段")
        self.compare_log = QListWidget()

        left_button = QPushButton("选择Excel A")
        right_button = QPushButton("选择Excel B")
        left_button.clicked.connect(lambda: self._choose_compare_file(self.compare_left_input))
        right_button.clicked.connect(lambda: self._choose_compare_file(self.compare_right_input))
        form = QFormLayout()
        form.addRow(left_button, self.compare_left_input)
        form.addRow(right_button, self.compare_right_input)
        form.addRow("匹配主键", self.compare_key_input)
        run = QPushButton("开始对比")
        run.clicked.connect(self._run_compare)
        layout.addLayout(form)
        layout.addWidget(self.compare_field_selector, 1)
        layout.addWidget(QLabel("操作日志"))
        layout.addWidget(self.compare_log)
        layout.addWidget(run)
        return page

    def _chart_page(self) -> QWidget:
        """创建“图表生成”页面。

        当前图表页仍使用手动输入字段名；后续也可以复用 FieldSelector
        来自动填充 X/Y/分组字段候选。
        """

        page = QWidget()
        layout = self._page_layout(page, "图表生成器")
        self.chart_file_input = QLineEdit()
        self.chart_type = QComboBox()
        self.chart_type.addItems(["bar", "line", "pie"])
        self.chart_x_input = QLineEdit()
        self.chart_y_input = QLineEdit()
        self.chart_group_input = QLineEdit()
        self.chart_title_input = QLineEdit()

        pick = QPushButton("选择数据文件")
        pick.clicked.connect(lambda: self._choose_one_file(self.chart_file_input))
        form = QFormLayout()
        form.addRow(pick, self.chart_file_input)
        form.addRow("图表类型", self.chart_type)
        form.addRow("X轴/维度", self.chart_x_input)
        form.addRow("Y轴/指标", self.chart_y_input)
        form.addRow("分组字段", self.chart_group_input)
        form.addRow("标题", self.chart_title_input)
        run = QPushButton("生成PNG图表")
        run.clicked.connect(self._run_chart)
        layout.addLayout(form)
        layout.addStretch(1)
        layout.addWidget(run)
        return page

    def _history_page(self) -> QWidget:
        """创建历史任务页面，展示 `HistoryStore` 中的最近任务。"""

        page = QWidget()
        layout = self._page_layout(page, "历史任务")
        self.history_table = QTableWidget(0, 4)
        self.history_table.setHorizontalHeaderLabels(["时间", "任务", "状态", "说明"])
        layout.addWidget(self.history_table)
        self._refresh_history()
        return page

    def _page_layout(self, page: QWidget, title_text: str) -> QVBoxLayout:
        """页面通用布局工厂，统一边距和标题样式。"""

        layout = QVBoxLayout(page)
        layout.setContentsMargins(28, 24, 28, 24)
        title = QLabel(title_text)
        title.setObjectName("PageTitle")
        layout.addWidget(title)
        return layout

    def _choose_files(self, target: QListWidget) -> None:
        """通用多文件选择函数。

        上游：合并页调用。
        下游：文件路径写入 QListWidget，随后 `_choose_merge_files()` 触发字段解析。
        """

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择文件",
            str(Path.home()),
            "表格文件 (*.xlsx *.xls *.xlsm *.csv)",
        )
        target.clear()
        target.addItems(files)

    def _choose_merge_files(self) -> None:
        """合并页选择文件并解析字段。

        这是“上传文件 -> 后端解析 -> 前端展示字段选择器”的桌面版入口。
        """

        self._choose_files(self.merge_files)
        files = [self.merge_files.item(index).text() for index in range(self.merge_files.count())]
        self._parse_fields_for_selector(files, self.merge_field_selector, self.merge_log)

    def _choose_one_file(self, target: QLineEdit) -> None:
        """通用单文件选择函数，把路径写入指定输入框。"""

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择文件",
            str(Path.home()),
            "表格文件 (*.xlsx *.xls *.xlsm *.csv)",
        )
        if file_path:
            target.setText(file_path)

    def _choose_clean_file(self) -> None:
        """清洗页选择文件并解析字段。"""

        self._choose_one_file(self.clean_file_input)
        source = self.clean_file_input.text().strip()
        self._parse_fields_for_selector([source] if source else [], self.clean_field_selector, self.clean_log)

    def _choose_compare_file(self, target: QLineEdit) -> None:
        """对比页选择左/右文件并在两个文件都具备后刷新主键候选。"""

        self._choose_one_file(target)
        left = self.compare_left_input.text().strip()
        right = self.compare_right_input.text().strip()
        files = [path for path in [left, right] if path]
        self._parse_fields_for_selector(files, self.compare_field_selector, self.compare_log)
        if len(files) == 2:
            # 主键只能从两个文件都存在的字段中选择，减少执行时缺字段失败。
            common_fields = [
                field.name
                for field in self.compare_field_selector.fields
                if field.files == field.total_files
            ]
            self.compare_key_input.clear()
            self.compare_key_input.addItems(common_fields)

    def _parse_fields_for_selector(self, files: list[str], selector: FieldSelector, log: QListWidget) -> None:
        """解析文件字段并刷新字段选择器。

        上游：合并、清洗、对比页的文件选择事件。
        下游：
        - `parse_files_metadata()` 读取表头和 dtype。
        - `build_field_coverage()` 统计公共字段/部分字段。
        - `selector.set_fields()` 渲染 UI。
        """

        if not files:
            return
        try:
            self._append_log(log, f"开始解析 {len(files)} 个文件的表头")
            metadata = parse_files_metadata(files)
            fields = build_field_coverage(metadata)
            selector.set_fields(fields)
            common_count = sum(1 for field in fields if field.is_common)
            self._append_log(log, f"解析完成：共 {len(fields)} 个字段，公共字段 {common_count} 个")
        except Exception as exc:
            selector.set_fields([])
            self._append_log(log, f"解析失败：{exc}")
            self._warn(str(exc))

    def _save_path(self, title: str, suffix: str) -> Path | None:
        """弹出保存文件对话框，返回用户选择的输出路径。"""

        path, _ = QFileDialog.getSaveFileName(self, title, str(Path.home() / f"处理结果{suffix}"), f"*{suffix}")
        return Path(path) if path else None

    def _run_merge(self) -> None:
        """执行合并任务。

        输入来源：
        - 文件列表：`self.merge_files`
        - 勾选字段：`self.merge_field_selector.selected_fields()`
        - 合并方式：单选按钮
        - 匹配字段：`self.merge_key_input`

        下游：调用 `ProcessService.merge_rows()` 或 `ProcessService.merge_by_key()`。
        """

        files = [self.merge_files.item(index).text() for index in range(self.merge_files.count())]
        if not files:
            self._warn("请先选择文件")
            return
        output = self._save_path("保存合并结果", ".xlsx")
        if not output:
            return
        try:
            selected_columns = self.merge_field_selector.selected_fields()
            if not selected_columns:
                self._warn("请至少选择一个字段")
                return
            self._append_log(self.merge_log, f"用户选择字段：{', '.join(selected_columns)}")
            if self.merge_key_radio.isChecked():
                key = self.merge_key_input.text().strip()
                if not key:
                    self._warn("请输入匹配字段")
                    return
                result = self.process_service.merge_by_key(files, key, output, columns=selected_columns)
                name = "字段匹配合并"
            else:
                result = self.process_service.merge_rows(files, output, columns=selected_columns)
                name = "纵向合并"
            self._append_log(self.merge_log, f"合并成功：{result}")
            self._record_success(name, files, [result], f"已导出：{result}")
        except Exception as exc:
            self._append_log(self.merge_log, f"合并失败：{exc}")
            self._record_failure("合并失败", files, str(exc))

    def _run_clean(self) -> None:
        """执行数据清洗任务。

        输入来源：
        - source：待清洗文件路径。
        - selected_columns：字段选择器中的勾选字段。
        - rename/drop/fill：页面表单输入。

        下游：组装 `CleanRule` 列表后传给 `ProcessService.clean_file()`。
        """

        source = self.clean_file_input.text().strip()
        if not source:
            self._warn("请先选择文件")
            return
        output = self._save_path("保存清洗结果", ".xlsx")
        if not output:
            return
        selected_columns = self.clean_field_selector.selected_fields()
        if not selected_columns:
            self._warn("请至少选择一个字段")
            return
        rules = [
            # 删除全空行不绑定字段，作用于整张表；其余规则按勾选字段执行。
            CleanRule(CleanAction.DROP_EMPTY_ROWS),
            CleanRule(CleanAction.DROP_DUPLICATES, columns=selected_columns),
            CleanRule(CleanAction.STRIP_TEXT, columns=selected_columns),
            CleanRule(CleanAction.NORMALIZE_DATES, columns=selected_columns),
            CleanRule(CleanAction.NORMALIZE_NUMBERS, columns=selected_columns),
            CleanRule(CleanAction.FILL_EMPTY, columns=selected_columns, fill_value=self.fill_value_input.text()),
        ]
        rename_map = self._parse_mapping(self.rename_input.text())
        if rename_map:
            # 改列名一般在清洗末尾执行，避免前面字段级规则找不到原字段。
            rules.append(CleanRule(CleanAction.RENAME_COLUMNS, rename_map=rename_map))
        drop_columns = self._parse_csv(self.drop_columns_input.text())
        if drop_columns:
            rules.append(CleanRule(CleanAction.DROP_COLUMNS, columns=drop_columns))
        try:
            self._append_log(self.clean_log, f"用户选择字段：{', '.join(selected_columns)}")
            result = self.process_service.clean_file(source, output, rules, columns=selected_columns)
            self._append_log(self.clean_log, f"清洗成功：{result}")
            self._record_success("数据清洗", [source], [result], f"已导出：{result}")
        except Exception as exc:
            self._append_log(self.clean_log, f"清洗失败：{exc}")
            self._record_failure("清洗失败", [source], str(exc))

    def _run_compare(self) -> None:
        """执行 Excel 对比任务。

        主键用于定位同一条业务数据；对比字段用于判断哪些列发生变化。
        两者分开，避免把“商品ID”这种定位字段也当成业务变化字段比较。
        """

        left = self.compare_left_input.text().strip()
        right = self.compare_right_input.text().strip()
        key = self.compare_key_input.currentText().strip()
        if not left or not right or not key:
            self._warn("请选择两个文件并填写对比字段")
            return
        output = self._save_path("保存对比结果", ".xlsx")
        if not output:
            return
        try:
            # 对比字段排除主键，因为主键只用于匹配同一行。
            selected_columns = [column for column in self.compare_field_selector.selected_fields() if column != key]
            if not selected_columns:
                self._warn("请至少选择一个非主键对比字段")
                return
            self._append_log(self.compare_log, f"匹配主键：{key}")
            self._append_log(self.compare_log, f"对比字段：{', '.join(selected_columns)}")
            result = self.process_service.compare_files(left, right, key, output, compare_columns=selected_columns)
            self._append_log(self.compare_log, f"对比成功：{result}")
            self._record_success("数据对比", [left, right], [result], f"已导出：{result}")
        except Exception as exc:
            self._append_log(self.compare_log, f"对比失败：{exc}")
            self._record_failure("对比失败", [left, right], str(exc))

    def _run_chart(self) -> None:
        """执行图表生成任务。

        输入来源：图表页的文件路径、图表类型、X/Y字段、可选分组字段。
        下游：`ChartService.create_chart()` 读取数据并生成 PNG。
        """

        source = self.chart_file_input.text().strip()
        x_column = self.chart_x_input.text().strip()
        y_column = self.chart_y_input.text().strip()
        if not source or not x_column or not y_column:
            self._warn("请选择文件并填写X轴、Y轴字段")
            return
        output = self._save_path("保存图表", ".png")
        if not output:
            return
        try:
            group = self.chart_group_input.text().strip() or None
            result = self.chart_service.create_chart(
                source,
                output,
                self.chart_type.currentText(),
                x_column,
                y_column,
                group,
                self.chart_title_input.text().strip(),
            )
            self._record_success("生成图表", [source], [result], f"已导出：{result}")
        except Exception as exc:
            self._record_failure("图表失败", [source], str(exc))

    def _record_success(self, name: str, inputs: list[str | Path], outputs: list[Path], message: str) -> None:
        """记录成功任务并刷新 UI。"""

        self.history.add(TaskRecord(name=name, status=TaskStatus.SUCCESS, input_files=[Path(path) for path in inputs], outputs=outputs, message=message))
        self._refresh_history()
        QMessageBox.information(self, "处理完成", message)

    def _record_failure(self, name: str, inputs: list[str | Path], message: str) -> None:
        """记录失败任务并弹出错误提示。"""

        self.history.add(TaskRecord(name=name, status=TaskStatus.FAILED, input_files=[Path(path) for path in inputs], message=message))
        self._refresh_history()
        QMessageBox.critical(self, "处理失败", message)

    def _refresh_home_history(self) -> None:
        """刷新首页最近任务列表。"""

        if not hasattr(self, "home_history"):
            return
        self.home_history.clear()
        for item in self.history.list()[:8]:
            self.home_history.addItem(f"{item.get('created_at', '')}  {item.get('name', '')}  {item.get('status', '')}")

    def _refresh_history(self) -> None:
        """刷新历史任务页表格，并同步首页最近任务。"""

        if not hasattr(self, "history_table"):
            self._refresh_home_history()
            return
        items = self.history.list()
        self.history_table.setRowCount(len(items))
        for row, item in enumerate(items):
            values = [item.get("created_at", ""), item.get("name", ""), item.get("status", ""), item.get("message", "")]
            for column, value in enumerate(values):
                self.history_table.setItem(row, column, QTableWidgetItem(str(value)))
        self.history_table.resizeColumnsToContents()
        self._refresh_home_history()

    def _warn(self, message: str) -> None:
        """统一提示弹窗。"""

        QMessageBox.warning(self, "提示", message)

    def _append_log(self, widget: QListWidget, message: str) -> None:
        """向当前页面操作日志追加一行。"""

        widget.addItem(message)
        widget.scrollToBottom()

    def _parse_csv(self, value: str) -> list[str]:
        """解析逗号分隔输入，兼容中文逗号。"""

        return [item.strip() for item in value.replace("，", ",").split(",") if item.strip()]

    def _parse_mapping(self, value: str) -> dict[str, str]:
        """解析列名映射输入。

        支持两种写法：
        - 旧列名=新列名
        - 旧列名→新列名
        """

        mapping: dict[str, str] = {}
        for part in self._parse_csv(value):
            if "=" in part:
                old, new = part.split("=", 1)
            elif "→" in part:
                old, new = part.split("→", 1)
            else:
                continue
            if old.strip() and new.strip():
                mapping[old.strip()] = new.strip()
        return mapping

    def _apply_style(self) -> None:
        """集中设置应用样式。"""

        self.setStyleSheet(
            """
            QMainWindow, QWidget { background: #f7f8fb; color: #1f2937; font-size: 14px; }
            QListWidget { background: #ffffff; border: 0; padding: 10px; }
            QListWidget::item { padding: 12px 10px; border-radius: 6px; }
            QListWidget::item:selected { background: #e8f0ff; color: #174ea6; }
            QLabel#Title { font-size: 34px; font-weight: 700; }
            QLabel#Subtitle { color: #64748b; font-size: 16px; margin-bottom: 18px; }
            QLabel#PageTitle { font-size: 24px; font-weight: 700; margin-bottom: 12px; }
            QPushButton { background: #2563eb; color: #ffffff; border: 0; border-radius: 6px; padding: 10px 14px; }
            QPushButton:hover { background: #1d4ed8; }
            QPushButton#CardButton { background: #ffffff; color: #111827; border: 1px solid #e5e7eb; font-size: 18px; font-weight: 600; }
            QPushButton#CardButton:hover { border-color: #2563eb; color: #174ea6; }
            QLineEdit, QTextEdit, QComboBox { background: #ffffff; border: 1px solid #d1d5db; border-radius: 6px; padding: 8px; }
            QTableWidget { background: #ffffff; border: 1px solid #e5e7eb; }
            """
        )


def run() -> None:
    """创建 QApplication、显示主窗口并进入 Qt 事件循环。"""

    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
