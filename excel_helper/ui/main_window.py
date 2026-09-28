"""Excel批处理助手主窗口。

界面改造说明（对照《前端界面优化方案》）：
1. 骨架：左侧导航改为"图标 + 分组"，每个页面统一"页头（标题 + 说明 + 唯一主按钮）"，
   主按钮不再被挤到页面最底部。
2. 布局：合并/清洗/对比/图表页改成左右两栏，左栏配置、右栏结果，消灭单列堆叠带来的大面积留白。
3. 反馈：`QMessageBox` 模态弹窗改为常驻 `FeedbackBar`，成功时提供"打开文件"入口。
4. 日志：常驻 QListWidget 改为可折叠 `LogPanel`，默认收起，把页面空间还给字段与配置。
5. 文件：新增拖拽入口与单文件移除（`FileListWidget` / `SingleFilePicker`）。
6. 图表页：X/Y/分组由手输改为下拉（复用文件真实表头），生成后直接在界面内预览。
7. 历史页：新增搜索、打开输出文件、清空历史，状态用颜色区分。

分层关系保持不变：
    UI(MainWindow/FieldSelector/widgets) -> Service(ProcessService/ChartService) -> Core(pandas/openpyxl/matplotlib)
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QUrl
from PySide6.QtGui import QAction, QColor, QDesktopServices, QPixmap
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
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QStackedWidget,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_helper.core.field_parser import build_field_coverage, parse_files_metadata
from excel_helper.models.field import FieldCoverage, FileMetadata
from excel_helper.models.rule import CleanAction, CleanRule
from excel_helper.models.task import TaskRecord, TaskStatus
from excel_helper.services.chart_service import ChartService
from excel_helper.services.process_service import ProcessService
from excel_helper.storage.history import HistoryStore
from excel_helper.ui.field_selector import FieldSelector
from excel_helper.ui.theme import COLOR, SPACE, apply_theme
from excel_helper.ui.widgets import (
    EmptyState,
    FeedbackBar,
    FileListWidget,
    LogPanel,
    SectionCard,
    SingleFilePicker,
)

# 左侧导航结构：(类型, 文案, 页面索引, 图标)
# 类型 item 可点击并切换到对应页面；类型 group 只是分组标题（不可选中）。
NAV_STRUCTURE: list[tuple[str, str, int | None, QStyle.StandardPixmap | None]] = [
    ("item", "首页", 0, QStyle.StandardPixmap.SP_ComputerIcon),
    ("group", "数据处理", None, None),
    ("item", "批量合并", 1, QStyle.StandardPixmap.SP_FileDialogDetailedView),
    ("item", "数据清洗", 2, QStyle.StandardPixmap.SP_DialogResetButton),
    ("item", "数据对比", 3, QStyle.StandardPixmap.SP_FileDialogListView),
    ("group", "输出与记录", None, None),
    ("item", "图表生成", 4, QStyle.StandardPixmap.SP_FileDialogInfoView),
    ("item", "历史任务", 5, QStyle.StandardPixmap.SP_BrowserReload),
]

# 首页功能入口卡：(标题, 说明, 目标页面索引)
HOME_FUNCTION_CARDS: list[tuple[str, str, int]] = [
    ("批量合并 Excel", "多文件纵向合并，或按主键字段横向匹配", 1),
    ("数据清洗", "去空行 / 去重 / 文本与日期标准化", 2),
    ("数据对比", "按主键输出新增、删除、修改三类差异", 3),
    ("生成图表", "柱状图 / 折线图 / 饼图，导出 PNG", 4),
]

# 清洗规则开关：(规则键, 界面文案)，默认全部勾选，等价于改造前的固定规则链。
CLEAN_RULE_OPTIONS: list[tuple[str, str]] = [
    (CleanAction.DROP_EMPTY_ROWS.value, "删除空行（整行为空）"),
    (CleanAction.DROP_DUPLICATES.value, "删除重复行（按勾选字段判断）"),
    (CleanAction.STRIP_TEXT.value, "文本去空格（含表头字段名）"),
    (CleanAction.NORMALIZE_DATES.value, "日期标准化为 YYYY-MM-DD"),
    (CleanAction.NORMALIZE_NUMBERS.value, "数字标准化（￥1,200元 → 1200）"),
    (CleanAction.FILL_EMPTY.value, "空值填充"),
    (CleanAction.RENAME_COLUMNS.value, "修改列名"),
    (CleanAction.DROP_COLUMNS.value, "删除列"),
]

# 图表类型下拉项：(显示文案, 传给 chart_factory 的值)
CHART_TYPE_OPTIONS: list[tuple[str, str]] = [("柱状图", "bar"), ("折线图", "line"), ("饼图", "pie")]

NO_GROUP_LABEL = "（不分组）"


class MainWindow(QMainWindow):
    """主窗口：负责页面编排、收集用户输入，并把处理任务交给 service 层。

    主窗口负责三件事：
    1. 创建页面和控件，收集用户输入。
    2. 在用户选择文件后调用 `field_parser` 解析字段，并刷新 `FieldSelector`。
    3. 在用户点击执行按钮后调用 service 层完成处理、导出和历史记录。
    """

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Excel批处理助手")
        self.resize(1280, 800)
        self.setMinimumSize(1080, 700)

        # service 层封装“读取 -> 处理 -> 导出”，UI 不直接写 pandas 逻辑。
        self.process_service = ProcessService()
        self.chart_service = ChartService()

        # history 负责本地 JSON 历史记录，供首页和历史任务页展示。
        self.history = HistoryStore()

        # 应用设计令牌与全局样式表（theme.py 是唯一样式入口）。
        apply_theme(self)

        # 左侧导航 + 右侧堆叠页面，是整个桌面端的主布局。
        self.stack = QStackedWidget()
        # 导航里包含不可点击的分组标题，因此单独维护"导航行号 -> 页面索引"映射。
        self._nav_to_stack: dict[int, int] = {}
        self._build_nav()

        self.stack.addWidget(self._home_page())
        self.stack.addWidget(self._merge_page())
        self.stack.addWidget(self._clean_page())
        self.stack.addWidget(self._compare_page())
        self.stack.addWidget(self._chart_page())
        self.stack.addWidget(self._history_page())

        # 底部反馈条：全局唯一，替代改造前各页面的模态弹窗。
        self.feedback = FeedbackBar()
        self.feedback.show_idle("就绪 · 选择文件后会自动解析表头与字段覆盖率")

        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)
        content = QHBoxLayout()
        content.setContentsMargins(0, 0, 0, 0)
        content.setSpacing(0)
        content.addWidget(self.nav)
        content.addWidget(self.stack, 1)
        root_layout.addLayout(content, 1)
        root_layout.addWidget(self.feedback)
        self.setCentralWidget(root)

        self._build_menubar()
        self._select_page(0)
        self._refresh_history()

    # ==================================================================
    # 骨架：导航 / 菜单 / 页头
    # ==================================================================

    def _build_nav(self) -> None:
        """构建左侧导航。

        改造前是 6 行等权重的文字项；现在加了图标与分组标题，
        并用 `_nav_to_stack` 维护行号到页面索引的映射，避免分组标题打乱索引。
        """

        self.nav = QListWidget()
        self.nav.setObjectName("Nav")
        self.nav.setFixedWidth(204)
        self.nav.setFrameShape(QFrame.Shape.NoFrame)

        for kind, label, stack_index, icon in NAV_STRUCTURE:
            item = QListWidgetItem(label)
            if kind == "group":
                # 分组标题：去掉所有交互标志，只作为视觉分隔，QSS 用 :disabled 弱化显示。
                item.setFlags(Qt.ItemFlag.NoItemFlags)
            else:
                assert stack_index is not None and icon is not None
                item.setIcon(self.style().standardIcon(icon))
                self._nav_to_stack[self.nav.count()] = stack_index
            self.nav.addItem(item)

        self.nav.currentRowChanged.connect(self._on_nav_row_changed)

    def _on_nav_row_changed(self, row: int) -> None:
        """导航行变化时切换页面；点到分组标题（无对应页面）时直接忽略。"""

        stack_index = self._nav_to_stack.get(row)
        if stack_index is None:
            return
        self.stack.setCurrentIndex(stack_index)

    def _select_page(self, stack_index: int) -> None:
        """按页面索引选中导航项（首页功能卡与"查看全部"走这里）。"""

        for row, index in self._nav_to_stack.items():
            if index == stack_index:
                self.nav.setCurrentRow(row)
                return

    def _build_menubar(self) -> None:
        """菜单栏只保留低频操作，避免和页面主按钮抢注意力。"""

        view_menu = self.menuBar().addMenu("视图")
        refresh_action = QAction("刷新历史", self)
        refresh_action.triggered.connect(self._refresh_history)
        view_menu.addAction(refresh_action)

        help_menu = self.menuBar().addMenu("帮助")
        about_action = QAction("关于", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    def _show_about(self) -> None:
        """关于对话框（低频信息，保留模态展示）。"""

        QMessageBox.about(
            self,
            "关于 Excel批处理助手",
            "Excel批处理助手\n\n覆盖批量合并、数据清洗、数据对比与图表生成的桌面小工具。\n"
            "技术栈：Python + PySide6 + pandas + matplotlib",
        )

    def _page_header(
        self,
        title: str,
        description: str,
        action_text: str | None = None,
        action_slot=None,
    ) -> tuple[QWidget, QPushButton | None]:
        """页面统一页头：左侧标题+说明，右侧唯一主按钮。

        改造前主按钮在页面最底部，长表单下用户要滚动才能执行；
        现在固定在页头右侧，始终在视线范围内。
        """

        header = QWidget()
        header.setObjectName("PageHeader")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(SPACE["m"])

        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        title_label = QLabel(title)
        title_label.setObjectName("PageTitle")
        desc_label = QLabel(description)
        desc_label.setObjectName("PageDesc")
        text_box.addWidget(title_label)
        text_box.addWidget(desc_label)
        layout.addLayout(text_box, 1)

        button: QPushButton | None = None
        if action_text:
            button = QPushButton(action_text)
            button.setObjectName("BtnPrimary")
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            if action_slot is not None:
                button.clicked.connect(action_slot)
            layout.addWidget(button)
        return header, button

    @staticmethod
    def _chip(text: str) -> QLabel:
        """创建一个浅蓝色信息徽章（用于"共 N 个文件""已选 N 个字段"）。"""

        chip = QLabel(text)
        chip.setObjectName("ChipInfo")
        return chip

    @staticmethod
    def _link_button(text: str, slot) -> QPushButton:
        """创建无边框文字按钮。"""

        button = QPushButton(text)
        button.setObjectName("BtnLink")
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        button.clicked.connect(slot)
        return button

    @staticmethod
    def _page_container(margins: int = SPACE["l"]) -> tuple[QWidget, QVBoxLayout]:
        """创建统一边距的页面容器。"""

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(margins + SPACE["xs"], margins, margins + SPACE["xs"], margins)
        layout.setSpacing(SPACE["m"])
        return page, layout

    @staticmethod
    def _form_label(text: str) -> QLabel:
        """表单行标签（统一弱化色，避免和内容抢注意力）。"""

        label = QLabel(text)
        label.setObjectName("Muted")
        return label

    @staticmethod
    def _form_layout() -> QFormLayout:
        """统一间距的表单布局。"""

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setHorizontalSpacing(SPACE["s"] + 4)
        form.setVerticalSpacing(SPACE["s"] + 2)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        return form

    # ==================================================================
    # 首页
    # ==================================================================

    def _home_page(self) -> QWidget:
        """创建首页：统计概览 + 功能入口 + 最近任务。"""

        page, layout = self._page_container(SPACE["xl"] - 8)

        header, _ = self._page_header("Excel批处理助手", "拖入文件，批量处理，自动对比，生成图表")
        layout.addWidget(header)

        # 统计条：数据直接聚合自本地历史记录，不引入新的状态源。
        layout.addLayout(self._build_stat_row())

        # 功能入口卡：点击整张卡切换到对应页面。
        cards = QHBoxLayout()
        cards.setSpacing(SPACE["m"] - 2)
        for title, description, stack_index in HOME_FUNCTION_CARDS:
            cards.addWidget(self._function_card(title, description, stack_index))
        layout.addLayout(cards)

        # 最近任务：改造前是纯文本行，现在是有状态色的表格 + 打开入口。
        recent_card = SectionCard("最近任务", extra=self._link_button("查看全部 →", lambda: self._select_page(5)))
        self.home_history = self._build_record_table(with_action=True)
        recent_card.add(self.home_history, 1)
        layout.addWidget(recent_card, 1)

        self._refresh_home_history()
        return page

    def _build_stat_row(self) -> QHBoxLayout:
        """构建首页统计条，并保存数值标签引用供刷新时更新。"""

        row = QHBoxLayout()
        row.setSpacing(SPACE["m"] - 2)
        # key -> 数值标签，刷新时只改文本，不重建控件。
        self.stat_labels: dict[str, QLabel] = {}
        for key, caption in [
            ("today", "今日任务"),
            ("success", "今日成功"),
            ("failed", "今日失败"),
            ("outputs", "累计输出文件"),
        ]:
            card = QFrame()
            card.setObjectName("Card")
            box = QVBoxLayout(card)
            box.setContentsMargins(SPACE["m"], SPACE["m"] - 4, SPACE["m"], SPACE["m"] - 4)
            box.setSpacing(2)
            value = QLabel("0")
            value.setObjectName("StatValue")
            caption_label = QLabel(caption)
            caption_label.setObjectName("StatKey")
            box.addWidget(value)
            box.addWidget(caption_label)
            self.stat_labels[key] = value
            row.addWidget(card)
        return row

    def _function_card(self, title: str, description: str, stack_index: int) -> QWidget:
        """功能入口卡。

        用 QPushButton 承载卡片样式，点击、键盘焦点、hover 都是原生行为，
        不需要额外写鼠标事件。
        """

        card = QPushButton()
        card.setObjectName("FunctionCard")
        card.setMinimumHeight(104)
        card.setCursor(Qt.CursorShape.PointingHandCursor)
        card.clicked.connect(lambda _checked=False, index=stack_index: self._select_page(index))

        layout = QVBoxLayout(card)
        layout.setContentsMargins(SPACE["m"], SPACE["m"] - 2, SPACE["m"], SPACE["m"] - 2)
        layout.setSpacing(4)
        title_label = QLabel(title)
        title_label.setObjectName("FunctionTitle")
        desc_label = QLabel(description)
        desc_label.setObjectName("MutedSmall")
        desc_label.setWordWrap(True)
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addStretch(1)
        return card

    # ==================================================================
    # 批量合并
    # ==================================================================

    def _merge_page(self) -> QWidget:
        """创建“批量合并”页面。

        页面数据流：
        选择/拖入文件 -> 解析表头 -> FieldSelector 选字段 -> 选择合并方式 -> ProcessService 导出。
        布局：左栏（文件 + 合并方式），右栏（字段选择），底部折叠日志。
        """

        page, layout = self._page_container()
        header, self.merge_run_button = self._page_header(
            "批量合并 Excel",
            "多个表格纵向拼接，或按主键字段横向匹配",
            "开始合并 ▸",
            self._run_merge,
        )
        self.merge_run_button.setEnabled(False)
        layout.addWidget(header)

        columns = QHBoxLayout()
        columns.setSpacing(SPACE["m"])

        # ---- 左栏：文件 + 合并方式 ----
        left = QVBoxLayout()
        left.setSpacing(SPACE["m"])

        self.merge_file_chip = self._chip("共 0 个文件")
        files_card = SectionCard("选择文件", step=1, extra=self.merge_file_chip)
        self.merge_files = FileListWidget()
        self.merge_files.filesChanged.connect(self._on_merge_files_changed)
        files_card.add(self.merge_files)
        left.addWidget(files_card)

        mode_card = SectionCard("合并方式", step=2)
        self.merge_mode_group = QButtonGroup(page)
        self.merge_rows_radio = QRadioButton("按行纵向合并（上下拼接）")
        self.merge_key_radio = QRadioButton("按指定字段匹配（左右拼接）")
        self.merge_rows_radio.setChecked(True)
        self.merge_mode_group.addButton(self.merge_rows_radio)
        self.merge_mode_group.addButton(self.merge_key_radio)

        radio_row = QHBoxLayout()
        radio_row.setSpacing(SPACE["l"])
        radio_row.addWidget(self.merge_rows_radio)
        radio_row.addWidget(self.merge_key_radio)
        mode_card.body.addLayout(radio_row)

        key_row = QHBoxLayout()
        key_row.setSpacing(SPACE["s"])
        key_row.addWidget(self._form_label("匹配字段"))
        # 改造前是自由输入框，字段名拼错要到执行时才报错；现在只列公共字段。
        self.merge_key_input = QComboBox()
        self.merge_key_input.setEditable(True)
        self.merge_key_input.setEnabled(False)
        key_row.addWidget(self.merge_key_input, 1)
        mode_card.body.addLayout(key_row)
        # 只有选中"按字段匹配"时才启用匹配字段下拉。
        self.merge_key_radio.toggled.connect(self.merge_key_input.setEnabled)
        left.addWidget(mode_card)
        left.addStretch(1)

        # ---- 右栏：字段选择 ----
        right = QVBoxLayout()
        right.setSpacing(SPACE["m"])
        self.merge_selected_chip = self._chip("已选 0 个字段")
        fields_card = SectionCard("选择字段", step=3, extra=self.merge_selected_chip)
        self.merge_field_selector = FieldSelector("字段列表（公共字段默认勾选）")
        self.merge_field_selector.fieldsChanged.connect(self._on_merge_fields_changed)
        fields_card.add(self.merge_field_selector, 1)
        right.addWidget(fields_card, 1)

        columns.addLayout(left, 34)
        columns.addLayout(right, 66)
        layout.addLayout(columns, 1)

        self.merge_log = LogPanel("操作日志")
        layout.addWidget(self.merge_log)
        return page

    def _on_merge_files_changed(self, paths: list[str]) -> None:
        """合并页文件集合变化：更新统计、启用主按钮、重新解析字段。"""

        self.merge_file_chip.setText(f"共 {len(paths)} 个文件")
        self.merge_run_button.setEnabled(bool(paths))
        metadata, fields = self._load_fields(paths, self.merge_field_selector, self.merge_log)
        if metadata:
            # 把每个文件的字段数显示在文件卡片上，用户不必展开字段表就能判断文件是否选对。
            for item in metadata:
                self.merge_files.set_meta(str(item.path), f"{len(item.columns)} 个字段")
            self._refresh_merge_key_options(fields)

    def _on_merge_fields_changed(self, selected: list[str]) -> None:
        """字段勾选变化时更新徽章计数。"""

        self.merge_selected_chip.setText(f"已选 {len(selected)} 个字段")

    def _refresh_merge_key_options(self, fields: list[FieldCoverage]) -> None:
        """用公共字段刷新"匹配字段"下拉，并尽量保留用户原选择。"""

        common = [field.name for field in fields if field.is_common]
        current = self.merge_key_input.currentText()
        self.merge_key_input.clear()
        self.merge_key_input.addItems(common)
        if not common:
            self.merge_key_input.setEditText("")
            return
        if current in common:
            self.merge_key_input.setCurrentText(current)
        else:
            self.merge_key_input.setCurrentIndex(0)

    def _run_merge(self) -> None:
        """执行合并任务。

        输入来源：
        - 文件列表：`self.merge_files.paths()`
        - 勾选字段：`self.merge_field_selector.selected_fields()`
        - 合并方式：单选按钮
        - 匹配字段：`self.merge_key_input`

        下游：调用 `ProcessService.merge_rows()` 或 `ProcessService.merge_by_key()`。
        """

        files = self.merge_files.paths()
        if not files:
            self._warn("请先选择要合并的文件")
            return
        selected_columns = self.merge_field_selector.selected_fields()
        if not selected_columns:
            self._warn("请至少选择一个字段")
            return
        key = self.merge_key_input.currentText().strip()
        if self.merge_key_radio.isChecked() and not key:
            self._warn("请选择匹配字段")
            return
        output = self._save_path("保存合并结果", ".xlsx")
        if not output:
            return

        try:
            self._append_log(self.merge_log, f"用户选择字段：{', '.join(selected_columns)}")
            self._begin_task("正在合并，请稍候…")
            if self.merge_key_radio.isChecked():
                result = self.process_service.merge_by_key(files, key, output, columns=selected_columns)
                name = "字段匹配合并"
            else:
                result = self.process_service.merge_rows(files, output, columns=selected_columns)
                name = "纵向合并"
            self._append_log(self.merge_log, f"合并成功：{result}")
            self._record_success(name, files, [result], f"合并完成，已导出：{result}")
        except Exception as exc:
            # 单次任务失败不影响应用稳定性：写日志 + 反馈条报错 + 记入历史。
            self._append_log(self.merge_log, f"合并失败：{exc}")
            self._record_failure("合并失败", files, str(exc))

    # ==================================================================
    # 数据清洗
    # ==================================================================

    def _clean_page(self) -> QWidget:
        """创建“数据清洗”页面。

        页面数据流：
        选择文件 -> 解析表头 -> 勾选清洗规则 + 选择字段 -> 组装 CleanRule -> ProcessService 导出。
        改造点：清洗规则从"硬编码不可见"改为"可勾选的规则清单"（`CleanRule.enabled` 已支持）。
        """

        page, layout = self._page_container()
        header, self.clean_run_button = self._page_header(
            "批量数据清洗",
            "去空行、去重、文本与日期标准化，按字段执行",
            "开始清洗 ▸",
            self._run_clean,
        )
        self.clean_run_button.setEnabled(False)
        layout.addWidget(header)

        columns = QHBoxLayout()
        columns.setSpacing(SPACE["m"])

        # ---- 左栏：文件 + 规则 ----
        left = QVBoxLayout()
        left.setSpacing(SPACE["m"])

        file_card = SectionCard("待清洗文件", step=1)
        self.clean_picker = SingleFilePicker("未选择文件")
        self.clean_picker.fileChanged.connect(self._on_clean_file_changed)
        file_card.add(self.clean_picker)
        left.addWidget(file_card)

        rules_card = SectionCard("清洗规则", step=2, subtitle="默认全部启用，可按需关闭")
        self.clean_rule_checks: dict[str, QPushButton] = {}
        for rule_key, caption in CLEAN_RULE_OPTIONS:
            # 用可勾选的 QPushButton 表示规则开关：整行可点，比小方框更易点中。
            toggle = QPushButton(caption)
            toggle.setObjectName("RuleCheck")
            toggle.setCheckable(True)
            toggle.setChecked(True)
            toggle.setCursor(Qt.CursorShape.PointingHandCursor)
            self.clean_rule_checks[rule_key] = toggle
            rules_card.add(toggle)

        # 规则参数：保持原有输入方式，未填写的参数等价于不生效。
        param_form = self._form_layout()
        self.rename_input = QLineEdit()
        self.rename_input.setPlaceholderText("旧列名=新列名，多条用逗号分隔")
        self.drop_columns_input = QLineEdit()
        self.drop_columns_input.setPlaceholderText("成本价,内部备注")
        self.fill_value_input = QLineEdit("0")
        param_form.addRow(self._form_label("修改列名"), self.rename_input)
        param_form.addRow(self._form_label("删除列"), self.drop_columns_input)
        param_form.addRow(self._form_label("空值填充值"), self.fill_value_input)
        rules_card.body.addLayout(param_form)
        left.addWidget(rules_card)
        left.addStretch(1)

        # ---- 右栏：字段 + 日志 ----
        right = QVBoxLayout()
        right.setSpacing(SPACE["m"])
        self.clean_selected_chip = self._chip("已选 0 个字段")
        fields_card = SectionCard("选择需要清洗的字段", step=3, extra=self.clean_selected_chip)
        self.clean_field_selector = FieldSelector("字段列表")
        self.clean_field_selector.fieldsChanged.connect(
            lambda selected: self.clean_selected_chip.setText(f"已选 {len(selected)} 个字段")
        )
        fields_card.add(self.clean_field_selector, 1)
        right.addWidget(fields_card, 1)

        self.clean_log = LogPanel("操作日志")
        right.addWidget(self.clean_log)

        columns.addLayout(left, 36)
        columns.addLayout(right, 64)
        layout.addLayout(columns, 1)
        return page

    def _on_clean_file_changed(self, path: str) -> None:
        """清洗页文件变化：解析字段并启用主按钮。"""

        self.clean_run_button.setEnabled(bool(path))
        metadata, _ = self._load_fields([path] if path else [], self.clean_field_selector, self.clean_log)
        if metadata:
            item = metadata[0]
            self.clean_picker.set_meta(f"{item.sheet_name} · {len(item.columns)} 个字段")

    def _run_clean(self) -> None:
        """执行数据清洗任务。

        输入来源：
        - source：待清洗文件路径。
        - selected_columns：字段选择器中的勾选字段。
        - enabled：每条清洗规则的勾选状态（通过 `CleanRule.enabled` 透传给 core）。
        """

        source = self.clean_picker.path()
        if not source:
            self._warn("请先选择待清洗文件")
            return
        selected_columns = self.clean_field_selector.selected_fields()
        if not selected_columns:
            self._warn("请至少选择一个字段")
            return
        output = self._save_path("保存清洗结果", ".xlsx")
        if not output:
            return

        # 规则顺序与改造前完全一致，只把"是否启用"交给界面控制。
        def enabled(rule_key: str) -> bool:
            return self.clean_rule_checks[rule_key].isChecked()

        rules = [
            # 删除全空行不绑定字段，作用于整张表；其余规则按勾选字段执行。
            CleanRule(CleanAction.DROP_EMPTY_ROWS, enabled=enabled(CleanAction.DROP_EMPTY_ROWS.value)),
            CleanRule(
                CleanAction.DROP_DUPLICATES,
                enabled=enabled(CleanAction.DROP_DUPLICATES.value),
                columns=selected_columns,
            ),
            CleanRule(CleanAction.STRIP_TEXT, enabled=enabled(CleanAction.STRIP_TEXT.value), columns=selected_columns),
            CleanRule(
                CleanAction.NORMALIZE_DATES,
                enabled=enabled(CleanAction.NORMALIZE_DATES.value),
                columns=selected_columns,
            ),
            CleanRule(
                CleanAction.NORMALIZE_NUMBERS,
                enabled=enabled(CleanAction.NORMALIZE_NUMBERS.value),
                columns=selected_columns,
            ),
            CleanRule(
                CleanAction.FILL_EMPTY,
                enabled=enabled(CleanAction.FILL_EMPTY.value),
                columns=selected_columns,
                fill_value=self.fill_value_input.text(),
            ),
        ]
        # 改列名放在清洗末尾执行，避免前面字段级规则找不到原字段名。
        rules.append(
            CleanRule(
                CleanAction.RENAME_COLUMNS,
                enabled=enabled(CleanAction.RENAME_COLUMNS.value),
                rename_map=self._parse_mapping(self.rename_input.text()),
            )
        )
        rules.append(
            CleanRule(
                CleanAction.DROP_COLUMNS,
                enabled=enabled(CleanAction.DROP_COLUMNS.value),
                columns=self._parse_csv(self.drop_columns_input.text()),
            )
        )

        try:
            self._append_log(self.clean_log, f"用户选择字段：{', '.join(selected_columns)}")
            active_rules = [rule.action.value for rule in rules if rule.enabled]
            self._append_log(self.clean_log, f"启用规则 {len(active_rules)} 条：{', '.join(active_rules)}")
            self._begin_task("正在清洗，请稍候…")
            result = self.process_service.clean_file(source, output, rules, columns=selected_columns)
            self._append_log(self.clean_log, f"清洗成功：{result}")
            self._record_success("数据清洗", [source], [result], f"清洗完成，已导出：{result}")
        except Exception as exc:
            self._append_log(self.clean_log, f"清洗失败：{exc}")
            self._record_failure("清洗失败", [source], str(exc))

    # ==================================================================
    # 数据对比
    # ==================================================================

    def _compare_page(self) -> QWidget:
        """创建“数据对比”页面。

        页面数据流：
        选择 Excel A/B -> 解析公共字段 -> 选择匹配主键 -> 勾选对比字段 -> 导出差异报告。
        布局：左栏 A/B 文件卡（视觉上呈现"A 对比 B"），右栏主键 + 对比字段。
        """

        page, layout = self._page_container()
        header, self.compare_run_button = self._page_header(
            "Excel 数据对比",
            "按主键对比两表，输出新增 / 删除 / 修改三类差异",
            "开始对比 ▸",
            self._run_compare,
        )
        self.compare_run_button.setEnabled(False)
        layout.addWidget(header)

        columns = QHBoxLayout()
        columns.setSpacing(SPACE["m"])

        # ---- 左栏：两个文件 + 主键 ----
        left = QVBoxLayout()
        left.setSpacing(SPACE["m"])

        files_card = SectionCard("对比文件", step=1)
        self.compare_left_picker = SingleFilePicker("基准文件（A · 旧数据）")
        self.compare_right_picker = SingleFilePicker("对比文件（B · 新数据）")
        self.compare_left_picker.fileChanged.connect(self._on_compare_file_changed)
        self.compare_right_picker.fileChanged.connect(self._on_compare_file_changed)
        files_card.add(self.compare_left_picker)

        vs_label = QLabel("VS")
        vs_label.setObjectName("ChipInfo")
        vs_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        vs_label.setFixedWidth(44)
        vs_row = QHBoxLayout()
        vs_row.addStretch(1)
        vs_row.addWidget(vs_label)
        vs_row.addStretch(1)
        files_card.body.addLayout(vs_row)

        files_card.add(self.compare_right_picker)
        left.addWidget(files_card)

        key_card = SectionCard("匹配主键", step=2, subtitle="只能选择两个文件都存在的字段")
        self.compare_key_input = QComboBox()
        self.compare_key_input.setEditable(True)
        self.compare_key_input.setPlaceholderText("需先选择两个文件")
        key_card.add(self.compare_key_input)
        left.addWidget(key_card)
        left.addStretch(1)

        # ---- 右栏：对比字段 + 日志 ----
        right = QVBoxLayout()
        right.setSpacing(SPACE["m"])
        self.compare_selected_chip = self._chip("已选 0 个字段")
        fields_card = SectionCard("选择需要对比的字段", step=3, extra=self.compare_selected_chip)
        self.compare_field_selector = FieldSelector("字段列表（主键会自动排除）")
        self.compare_field_selector.fieldsChanged.connect(
            lambda selected: self.compare_selected_chip.setText(f"已选 {len(selected)} 个字段")
        )
        fields_card.add(self.compare_field_selector, 1)
        right.addWidget(fields_card, 1)

        self.compare_log = LogPanel("操作日志")
        right.addWidget(self.compare_log)

        columns.addLayout(left, 38)
        columns.addLayout(right, 62)
        layout.addLayout(columns, 1)
        return page

    def _on_compare_file_changed(self, _path: str) -> None:
        """对比页任一文件变化：重新解析字段并刷新主键候选。

        改造前主键候选只在选完两个文件后才刷新，容易出现"只解析了一半"的中间状态；
        现在任一文件变化都会重算，主键下拉与字段表始终一致。
        """

        left = self.compare_left_picker.path()
        right = self.compare_right_picker.path()
        files = [path for path in [left, right] if path]
        self.compare_run_button.setEnabled(len(files) == 2)

        metadata, fields = self._load_fields(files, self.compare_field_selector, self.compare_log)
        for item in metadata:
            # 把字段数回填到对应的文件卡上。
            target = self.compare_left_picker if str(item.path) == left else self.compare_right_picker
            target.set_meta(f"{item.sheet_name} · {len(item.columns)} 个字段")

        if len(files) < 2:
            self.compare_key_input.clear()
            return
        # 主键只能从两个文件都存在的字段中选择，减少执行时缺字段失败。
        common_fields = [field.name for field in fields if field.is_common]
        current = self.compare_key_input.currentText()
        self.compare_key_input.clear()
        self.compare_key_input.addItems(common_fields)
        if current in common_fields:
            self.compare_key_input.setCurrentText(current)

    def _run_compare(self) -> None:
        """执行 Excel 对比任务。

        主键用于定位同一条业务数据；对比字段用于判断哪些列发生变化。
        两者分开，避免把“商品ID”这种定位字段也当成业务变化字段比较。
        """

        left = self.compare_left_picker.path()
        right = self.compare_right_picker.path()
        key = self.compare_key_input.currentText().strip()
        if not left or not right:
            self._warn("请选择两个待对比文件")
            return
        if not key:
            self._warn("请选择匹配主键")
            return
        # 对比字段排除主键，因为主键只用于匹配同一行。
        selected_columns = [column for column in self.compare_field_selector.selected_fields() if column != key]
        if not selected_columns:
            self._warn("请至少选择一个非主键对比字段")
            return
        output = self._save_path("保存对比结果", ".xlsx")
        if not output:
            return

        try:
            self._append_log(self.compare_log, f"匹配主键：{key}")
            self._append_log(self.compare_log, f"对比字段：{', '.join(selected_columns)}")
            self._begin_task("正在对比，请稍候…")
            result = self.process_service.compare_files(left, right, key, output, compare_columns=selected_columns)
            self._append_log(self.compare_log, f"对比成功：{result}")
            self._record_success("数据对比", [left, right], [result], f"对比完成，已导出：{result}")
        except Exception as exc:
            self._append_log(self.compare_log, f"对比失败：{exc}")
            self._record_failure("对比失败", [left, right], str(exc))

    # ==================================================================
    # 图表生成
    # ==================================================================

    def _chart_page(self) -> QWidget:
        """创建“图表生成”页面。

        改造点（对应迭代计划 S1-03）：
        - X/Y/分组字段由手输改为下拉，候选项来自所选文件的真实表头。
        - 生成结果直接在右侧卡片内预览，不必先导出再打开文件。
        """

        page, layout = self._page_container()
        header, self.chart_run_button = self._page_header(
            "图表生成器",
            "从文件真实表头中选择字段，生成柱状图 / 折线图 / 饼图",
            "生成图表 ▸",
            self._run_chart,
        )
        self.chart_run_button.setEnabled(False)
        layout.addWidget(header)

        columns = QHBoxLayout()
        columns.setSpacing(SPACE["m"])

        # ---- 左栏：文件 + 配置 ----
        left = QVBoxLayout()
        left.setSpacing(SPACE["m"])

        file_card = SectionCard("数据文件", step=1)
        self.chart_picker = SingleFilePicker("未选择文件")
        self.chart_picker.fileChanged.connect(self._on_chart_file_changed)
        file_card.add(self.chart_picker)
        left.addWidget(file_card)

        config_card = SectionCard("图表配置", step=2)
        # 当前图表文件的字段类型，供执行前校验 Y 轴是否为数值列。
        self.chart_columns = []
        self.chart_type = QComboBox()
        for caption, value in CHART_TYPE_OPTIONS:
            self.chart_type.addItem(caption, value)
        self.chart_x_input = QComboBox()
        self.chart_y_input = QComboBox()
        self.chart_group_input = QComboBox()
        self.chart_title_input = QLineEdit()
        self.chart_title_input.setPlaceholderText("留空时自动使用「指标 按 维度 统计」")

        form = self._form_layout()
        form.addRow(self._form_label("图表类型"), self.chart_type)
        form.addRow(self._form_label("X轴 / 维度"), self.chart_x_input)
        form.addRow(self._form_label("Y轴 / 指标"), self.chart_y_input)
        form.addRow(self._form_label("分组字段"), self.chart_group_input)
        form.addRow(self._form_label("标题"), self.chart_title_input)
        config_card.body.addLayout(form)
        left.addWidget(config_card)
        left.addStretch(1)

        # ---- 右栏：预览 + 日志 ----
        right = QVBoxLayout()
        right.setSpacing(SPACE["m"])
        preview_card = SectionCard("图表预览")
        self.chart_preview = QLabel()
        self.chart_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.chart_preview.setMinimumHeight(300)
        self.chart_preview.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        self.chart_empty = EmptyState("还没有生成图表", "选择文件并配置字段后，点击右上角「生成图表」")
        preview_card.add(self.chart_preview, 1)
        preview_card.add(self.chart_empty, 1)
        self.chart_preview.setVisible(False)
        right.addWidget(preview_card, 1)

        self.chart_log = LogPanel("操作日志")
        right.addWidget(self.chart_log)

        columns.addLayout(left, 42)
        columns.addLayout(right, 58)
        layout.addLayout(columns, 1)
        return page

    def _on_chart_file_changed(self, path: str) -> None:
        """图表页文件变化：解析表头并填充三个字段下拉。"""

        self.chart_run_button.setEnabled(bool(path))
        if not path:
            self.chart_columns = []
            for combo in (self.chart_x_input, self.chart_y_input, self.chart_group_input):
                combo.clear()
            return
        metadata, _ = self._load_fields([path], None, self.chart_log)
        if not metadata:
            return
        item = metadata[0]
        self.chart_columns = list(item.columns)
        self.chart_picker.set_meta(f"{item.sheet_name} · {len(item.columns)} 个字段")

        names = [column.name for column in item.columns]
        numeric_names = [column.name for column in item.columns if self._is_numeric_dtype(column.dtype)]
        self._fill_combo(self.chart_x_input, names)
        # Y 轴必须是数值字段：默认取第一个数值列，没有数值列时退回第一个字段（执行前还会再校验一次）。
        self._fill_combo(self.chart_y_input, names, default_name=numeric_names[0] if numeric_names else None)
        # 分组字段可留空，第一项明确表示“不分组”。
        self.chart_group_input.clear()
        self.chart_group_input.addItem(NO_GROUP_LABEL, None)
        for name in names:
            self.chart_group_input.addItem(name, name)

    @staticmethod
    def _is_numeric_dtype(dtype: str) -> bool:
        """判断字段类型是否可参与求和。

        `chart_factory` 内部使用 `groupby(...).sum()`，只接受数值列；
        这里提前判断，避免用户选到文本列后才在执行阶段收到晦涩的 pandas 报错。
        """

        return bool(re.search(r"(int|float|double|number)", dtype, re.IGNORECASE))

    def _chart_column_dtype(self, name: str) -> str:
        """查询当前图表文件的字段类型，用于执行前校验。"""

        for column in getattr(self, "chart_columns", []):
            if column.name == name:
                return column.dtype
        return ""

    @staticmethod
    def _fill_combo(combo: QComboBox, items: list[str], default_name: str | None = None) -> None:
        """用字段名填充下拉框，并尽量保持用户原选择。

        参数 default_name：没有历史选择时优先选中的字段（例如图表 Y 轴优先选第一个数值列）。
        """

        current = combo.currentText()
        combo.clear()
        combo.addItems(items)
        if current in items:
            combo.setCurrentText(current)
        elif default_name in items:
            combo.setCurrentText(str(default_name))
        elif items:
            combo.setCurrentIndex(0)

    def _run_chart(self) -> None:
        """执行图表生成任务。

        输入来源：文件路径、图表类型、X/Y 字段、可选分组字段、标题。
        下游：`ChartService.create_chart()` 读取数据并生成 PNG，然后回到界面内预览。
        """

        source = self.chart_picker.path()
        x_column = self.chart_x_input.currentText().strip()
        y_column = self.chart_y_input.currentText().strip()
        if not source or not x_column or not y_column:
            self._warn("请选择文件并指定 X 轴、Y 轴字段")
            return
        # Y 轴参与 groupby().sum()，必须是数值列；提前拦截并给出可操作的提示，
        # 而不是把 pandas 的原始报错抛给用户。
        y_dtype = self._chart_column_dtype(y_column)
        if y_dtype and not self._is_numeric_dtype(y_dtype):
            self._warn(f"Y 轴需要选择数值字段，当前「{y_column}」是文本列（{y_dtype}），请重新选择")
            return
        output = self._save_path("保存图表", ".png")
        if not output:
            return

        try:
            group = self.chart_group_input.currentData()
            self._append_log(
                self.chart_log,
                f"图表类型：{self.chart_type.currentText()}，X={x_column}，Y={y_column}",
            )
            self._begin_task("正在生成图表，请稍候…")
            result = self.chart_service.create_chart(
                source,
                output,
                self.chart_type.currentData(),
                x_column,
                y_column,
                group,
                self.chart_title_input.text().strip(),
            )
            self._append_log(self.chart_log, f"图表已生成：{result}")
            self._show_chart_preview(result)
            self._record_success("生成图表", [source], [result], f"图表已生成：{result}")
        except Exception as exc:
            self._append_log(self.chart_log, f"图表生成失败：{exc}")
            self._record_failure("图表失败", [source], str(exc))

    def _show_chart_preview(self, path: Path) -> None:
        """把生成的 PNG 显示在预览区，并按控件尺寸等比缩放。"""

        pixmap = QPixmap(str(path))
        if pixmap.isNull():
            return
        target = self.chart_preview.size()
        if target.width() < 50 or target.height() < 50:
            # 控件还没完成布局时退回到固定尺寸，避免缩放成 0 像素。
            target = QSize(640, 320)
        self.chart_preview.setPixmap(
            pixmap.scaled(target, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        )
        self.chart_empty.setVisible(False)
        self.chart_preview.setVisible(True)

    # ==================================================================
    # 历史任务
    # ==================================================================

    def _history_page(self) -> QWidget:
        """创建历史任务页面，展示 `HistoryStore` 中的最近任务。

        改造点（对应迭代计划 S3-04 的一部分）：新增搜索、打开输出文件、清空历史。
        "重跑任务"需要保存完整配置快照，放在后续迭代实现。
        """

        page, layout = self._page_container()
        header, _ = self._page_header("历史任务", "记录每一次处理的结果，可直接打开输出文件")
        layout.addWidget(header)

        toolbar = QWidget()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        toolbar_layout.setSpacing(SPACE["s"])
        self.history_search = QLineEdit()
        self.history_search.setPlaceholderText("搜索任务名 / 说明")
        self.history_search.setClearButtonEnabled(True)
        self.history_search.textChanged.connect(self._refresh_history)
        toolbar_layout.addWidget(self.history_search, 1)

        refresh_button = QPushButton("刷新")
        refresh_button.setObjectName("BtnSecondary")
        refresh_button.setCursor(Qt.CursorShape.PointingHandCursor)
        refresh_button.clicked.connect(self._refresh_history)
        toolbar_layout.addWidget(refresh_button)

        clear_button = QPushButton("清空历史")
        clear_button.setObjectName("BtnDanger")
        clear_button.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_button.clicked.connect(self._clear_history)
        toolbar_layout.addWidget(clear_button)

        card = SectionCard("任务记录", extra=toolbar)
        self.history_table = self._build_record_table(with_action=True)
        self.history_empty = EmptyState("还没有任务记录", "完成一次合并 / 清洗 / 对比 / 图表后，记录会出现在这里")
        card.add(self.history_table, 1)
        card.add(self.history_empty, 1)
        layout.addWidget(card, 1)
        return page

    def _clear_history(self) -> None:
        """清空历史记录（不可恢复操作，先二次确认）。"""

        if not self.history.list():
            self._warn("当前没有历史记录")
            return
        confirm = QMessageBox.question(
            self,
            "清空历史",
            "确定要清空全部历史任务记录吗？该操作不可恢复。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.history.clear()
        self._refresh_history()
        self.feedback.show_success("历史记录已清空")

    # ==================================================================
    # 数据解析与任务反馈
    # ==================================================================

    def _load_fields(
        self,
        files: list[str],
        selector: FieldSelector | None,
        log: LogPanel | None,
    ) -> tuple[list[FileMetadata], list[FieldCoverage]]:
        """解析文件表头并按需刷新字段选择器。

        上游：合并、清洗、对比、图表页的文件选择事件。
        下游：
        - `parse_files_metadata()` 读取表头和 dtype。
        - `build_field_coverage()` 统计公共字段 / 部分字段。
        - `selector.set_fields()` 渲染 UI（图表页不需要字段表，可传 None）。

        返回 (metadata, coverage)；解析失败时返回空列表并写日志 + 反馈条提示，不弹模态框。
        """

        if not files:
            if selector is not None:
                selector.set_fields([])
            return [], []
        try:
            if log is not None:
                self._append_log(log, f"开始解析 {len(files)} 个文件的表头")
            metadata = parse_files_metadata(files)
            fields = build_field_coverage(metadata)
            if selector is not None:
                selector.set_fields(fields)
            common_count = sum(1 for field in fields if field.is_common)
            if log is not None:
                self._append_log(log, f"解析完成：共 {len(fields)} 个字段，公共字段 {common_count} 个")
            return metadata, fields
        except Exception as exc:
            if selector is not None:
                selector.set_fields([])
            if log is not None:
                self._append_log(log, f"解析失败：{exc}")
            self._warn(f"文件解析失败：{exc}")
            return [], []

    def _begin_task(self, message: str) -> None:
        """任务开始：把反馈条切到“执行中”并强制重绘一次。

        当前 service 调用仍在主线程同步执行，这里先刷新界面让用户看到提示；
        真正的异步与可取消在迭代计划 S2-01 中落地。
        """

        self.feedback.show_busy(message)
        QApplication.processEvents()

    def _record_success(self, name: str, inputs: list[str | Path], outputs: list[Path], message: str) -> None:
        """记录成功任务，并在反馈条上提供“打开文件”入口。"""

        self.history.add(
            TaskRecord(
                name=name,
                status=TaskStatus.SUCCESS,
                input_files=[Path(path) for path in inputs],
                outputs=outputs,
                message=message,
            )
        )
        self._refresh_history()
        if outputs:
            first_output = outputs[0]
            self.feedback.show_success(message, "打开文件", lambda target=first_output: self._open_path(target))
        else:
            self.feedback.show_success(message)

    def _record_failure(self, name: str, inputs: list[str | Path], message: str) -> None:
        """记录失败任务并在反馈条上提示（不再弹模态框）。"""

        self.history.add(
            TaskRecord(
                name=name,
                status=TaskStatus.FAILED,
                input_files=[Path(path) for path in inputs],
                message=message,
            )
        )
        self._refresh_history()
        self.feedback.show_error(f"{name}：{message}")

    def _open_path(self, path: str | Path) -> None:
        """用系统默认程序打开输出文件。"""

        target = Path(path)
        if not target.exists():
            self.feedback.show_error(f"文件不存在或已被移动：{target}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))

    def _save_path(self, title: str, suffix: str) -> Path | None:
        """弹出保存文件对话框，返回用户选择的输出路径。"""

        path, _ = QFileDialog.getSaveFileName(self, title, str(Path.home() / f"处理结果{suffix}"), f"*{suffix}")
        return Path(path) if path else None

    def _warn(self, message: str) -> None:
        """统一校验提示入口。

        改造前这里弹 QMessageBox.warning，连续校验会连环弹窗；
        现在改为底部反馈条警告态，不打断操作流。
        """

        self.feedback.show_warning(message)

    def _append_log(self, widget: LogPanel, message: str) -> None:
        """向当前页面操作日志追加一行（LogPanel 保留了 QListWidget 的 addItem 接口）。"""

        widget.addItem(message)
        widget.scrollToBottom()

    # ==================================================================
    # 历史记录渲染
    # ==================================================================

    def _build_record_table(self, with_action: bool) -> QTableWidget:
        """创建任务记录表格（首页与历史页共用同一套列定义）。"""

        headers = ["时间", "任务", "状态", "说明"]
        if with_action:
            headers.append("操作")
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        table.setShowGrid(False)
        table.setWordWrap(False)
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        for column in range(len(headers) - 1):
            header.setSectionResizeMode(column, header.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(len(headers) - 1, header.ResizeMode.Stretch)
        return table

    def _refresh_home_history(self) -> None:
        """刷新首页最近任务（只展示最近 8 条）。"""

        if not hasattr(self, "home_history"):
            return
        self._fill_record_table(self.home_history, self.history.list()[:8])

    def _refresh_history(self) -> None:
        """刷新历史任务页表格、首页最近任务与统计条。"""

        items = self.history.list()
        if hasattr(self, "history_table"):
            keyword = self.history_search.text().strip().lower()
            filtered = [
                item
                for item in items
                if not keyword
                or keyword in str(item.get("name", "")).lower()
                or keyword in str(item.get("message", "")).lower()
            ]
            self._fill_record_table(self.history_table, filtered)
            # 无记录时显示空状态，避免用户面对一张空表。
            self.history_empty.setVisible(not filtered)
            self.history_table.setVisible(bool(filtered))
        self._refresh_home_history()
        self._refresh_stats(items)

    def _refresh_stats(self, items: list[dict[str, object]]) -> None:
        """刷新首页统计条：今日任务 / 今日成功 / 今日失败 / 累计输出文件。"""

        if not hasattr(self, "stat_labels"):
            return
        today = date.today().isoformat()
        today_items = [item for item in items if str(item.get("created_at", "")).startswith(today)]
        success = sum(1 for item in today_items if item.get("status") == TaskStatus.SUCCESS.value)
        failed = sum(1 for item in today_items if item.get("status") == TaskStatus.FAILED.value)
        outputs = sum(len(item.get("outputs") or []) for item in items)

        self.stat_labels["today"].setText(str(len(today_items)))
        self.stat_labels["success"].setText(str(success))
        self.stat_labels["success"].setStyleSheet(f"color:{COLOR['success']};")
        self.stat_labels["failed"].setText(str(failed))
        self.stat_labels["failed"].setStyleSheet(f"color:{COLOR['danger']};")
        self.stat_labels["outputs"].setText(str(outputs))

    def _fill_record_table(self, table: QTableWidget, records: list[dict[str, object]]) -> None:
        """把历史记录填充到表格。

        重建前必须先摘掉单元格里的按钮控件，否则旧的按钮会残留在视图上。
        """

        for row in range(table.rowCount()):
            for column in range(table.columnCount()):
                if table.cellWidget(row, column) is not None:
                    table.removeCellWidget(row, column)
        table.setRowCount(0)
        table.setRowCount(len(records))

        for row, item in enumerate(records):
            created_at = str(item.get("created_at", "")).replace("T", " ")
            values = [created_at, str(item.get("name", "")), "", str(item.get("message", ""))]
            for column, value in enumerate(values):
                cell = QTableWidgetItem(value)
                cell.setToolTip(value)
                table.setItem(row, column, cell)

            status_item = table.item(row, 2)
            if status_item is not None:
                succeeded = item.get("status") == TaskStatus.SUCCESS.value
                # 状态列用颜色区分：成功绿、失败红，替代改造前的一行纯文本。
                status_item.setText("成功" if succeeded else "失败")
                status_item.setForeground(QColor(COLOR["success"] if succeeded else COLOR["danger"]))
                font = status_item.font()
                font.setBold(True)
                status_item.setFont(font)

            if table.columnCount() > 4:
                outputs = [Path(path) for path in (item.get("outputs") or [])]
                if outputs:
                    table.setCellWidget(
                        row,
                        4,
                        self._link_button("打开输出", lambda _checked=False, target=outputs[0]: self._open_path(target)),
                    )

    # ==================================================================
    # 输入解析小工具
    # ==================================================================

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


def run() -> None:
    """创建 QApplication、显示主窗口并进入 Qt 事件循环。"""

    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
