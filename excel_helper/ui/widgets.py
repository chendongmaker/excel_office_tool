"""UI 通用组件库。

改造背景：改造前每个页面都在 `main_window.py` 里用 `QVBoxLayout` 直接堆控件，
导致样式、空状态、拖拽、反馈条这些重复逻辑散落各处。这里把可复用部分抽成组件，
让主窗口只负责"编排页面"，不再关心控件细节。

组件清单：
- `SectionCard`     卡片分组容器（步骤徽章 + 标题 + 右侧信息位）
- `EmptyState`      空状态占位（图标文案 + 操作入口）
- `FileListWidget`  多文件选择区（拖拽 + 选择 + 单文件移除 + 字段数元信息）
- `SingleFilePicker`单文件选择卡（拖拽 + 选择，用于清洗/对比/图表页）
- `LogPanel`        可折叠操作日志（对外保留 `addItem()`，兼容旧调用）
- `FeedbackBar`     页面底部反馈条（执行中/成功/失败/警告，替代模态弹窗）

层级约定：本模块只依赖 `theme.py` 与 `core.excel_reader` 的常量，
不依赖 `main_window.py`，避免循环导入。
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import QUrl, Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from excel_helper.core.excel_reader import TABLE_SUFFIXES
from excel_helper.ui.theme import COLOR, SPACE, refresh_style

# 文件对话框过滤器：与 `core.excel_reader.TABLE_SUFFIXES` 保持一致，避免两处定义漂移。
FILE_DIALOG_FILTER = "表格文件 (*.xlsx *.xls *.xlsm *.csv)"


def is_supported_file(path: str | Path) -> bool:
    """判断路径是否为支持处理的表格文件。"""

    return Path(path).suffix.lower() in TABLE_SUFFIXES


def collect_table_paths(urls: list[QUrl]) -> list[str]:
    """从拖拽内容中筛选出受支持的本地表格文件路径。

    参数 urls：拖拽事件 `mimeData().urls()` 的结果。
    返回：去重且保持原顺序的本地路径；非本地文件、目录、不支持的扩展名都会被剔除。
    """

    paths: list[str] = []
    for url in urls:
        local = url.toLocalFile()
        if not local:
            continue
        candidate = Path(local)
        if candidate.is_file() and is_supported_file(candidate) and str(candidate) not in paths:
            paths.append(str(candidate))
    return paths


def collect_rejected_table_paths(urls: list[QUrl]) -> list[str]:
    """收集拖入的本地文件中不支持的格式，供界面明确提示用户。"""

    rejected: list[str] = []
    for url in urls:
        local = url.toLocalFile()
        if not local:
            continue
        candidate = Path(local)
        if candidate.is_file() and not is_supported_file(candidate) and str(candidate) not in rejected:
            rejected.append(str(candidate))
    return rejected


# ---------------------------------------------------------------------------
# 卡片与空状态
# ---------------------------------------------------------------------------


class SectionCard(QFrame):
    """卡片分组容器。

    用法：
        card = SectionCard("选择文件", step=1, extra=QLabel("2 个文件"))
        card.add(my_widget)          # 往卡片主体追加控件
        card.body.addLayout(row)     # 或者直接往主体布局里加

    设计意图：把"这一步在做什么"用步骤徽章 + 标题显式表达，
    替代改造前"所有控件平铺在一个 QVBoxLayout 里"的堆叠式布局。
    """

    def __init__(
        self,
        title: str,
        step: int | None = None,
        subtitle: str = "",
        extra: QWidget | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("Card")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(SPACE["m"], SPACE["m"] - 2, SPACE["m"], SPACE["m"] - 2)
        outer.setSpacing(SPACE["s"])

        header = QHBoxLayout()
        header.setSpacing(SPACE["s"] - 2)
        if step is not None:
            badge = QLabel(str(step))
            badge.setObjectName("StepBadge")
            header.addWidget(badge)
        header.addWidget(self._label(title, "CardTitle"))
        header.addStretch(1)
        if extra is not None:
            header.addWidget(extra)
        outer.addLayout(header)

        if subtitle:
            outer.addWidget(self._label(subtitle, "MutedSmall"))

        # 主体布局：外部通过 add()/body 往里放业务控件。
        self.body = QVBoxLayout()
        self.body.setContentsMargins(0, 0, 0, 0)
        self.body.setSpacing(SPACE["s"])
        outer.addLayout(self.body)

    @staticmethod
    def _label(text: str, object_name: str) -> QLabel:
        """创建带 objectName 的标签，样式统一由 theme.py 的 QSS 控制。"""

        label = QLabel(text)
        label.setObjectName(object_name)
        return label

    def add(self, widget: QWidget, stretch: int = 0) -> None:
        """往卡片主体添加控件（stretch 用于让某个控件占满剩余高度）。"""

        self.body.addWidget(widget, stretch)


class EmptyState(QFrame):
    """空状态占位块。

    改造前：未选文件时字段表/历史表是一片空白，用户不知道下一步做什么。
    改造后：统一给出"当前没有什么 + 应该怎么开始"的引导。
    """

    def __init__(
        self,
        title: str,
        hint: str = "",
        actions: list[QPushButton] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("EmptyState")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(SPACE["l"], SPACE["l"], SPACE["l"], SPACE["l"])
        layout.setSpacing(SPACE["s"] - 2)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_label = QLabel(title)
        title_label.setObjectName("EmptyTitle")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        if hint:
            hint_label = QLabel(hint)
            hint_label.setObjectName("EmptyHint")
            hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(hint_label)

        if actions:
            row = QHBoxLayout()
            row.setSpacing(SPACE["s"])
            row.addStretch(1)
            for button in actions:
                row.addWidget(button)
            row.addStretch(1)
            layout.addLayout(row)


# ---------------------------------------------------------------------------
# 文件选择
# ---------------------------------------------------------------------------


class FileListWidget(QFrame):
    """多文件选择区（合并页使用）。

    对外契约：
    - `paths()`：当前文件路径列表，替代原先从 QListWidget 逐项取文本的写法。
    - `filesChanged` 信号：文件集合变化时发出，页面据此重新解析字段。
    - `set_meta(path, text)`：设置单个文件的辅助信息（例如"4 个字段"）。
    """

    filesChanged = Signal(list)
    # 将拖入的非法格式路径交给页面显示提示，避免静默忽略。
    filesRejected = Signal(list)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)

        self._paths: list[str] = []
        self._meta: dict[str, str] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(SPACE["m"] - 4, SPACE["m"] - 4, SPACE["m"] - 4, SPACE["m"] - 4)
        layout.setSpacing(SPACE["s"])

        hint = QLabel("把表格文件拖到这里，或点击下方按钮选择（支持多选）")
        hint.setObjectName("MutedSmall")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        # 文件行放在滚动区里，文件多时不会把页面撑高。
        self._rows_host = QWidget()
        self._rows_layout = QVBoxLayout(self._rows_host)
        self._rows_layout.setContentsMargins(0, 0, 0, 0)
        self._rows_layout.setSpacing(6)
        self._rows_layout.addStretch(1)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._scroll.setMaximumHeight(168)
        self._scroll.setWidget(self._rows_host)
        layout.addWidget(self._scroll)

        pick = QPushButton("选择文件")
        pick.setObjectName("BtnSecondary")
        pick.setCursor(Qt.CursorShape.PointingHandCursor)
        pick.clicked.connect(self._open_dialog)
        layout.addWidget(pick, alignment=Qt.AlignmentFlag.AlignLeft)

        self._empty_tip = QLabel("还没有选择文件")
        self._empty_tip.setObjectName("MutedSmall")
        layout.insertWidget(1, self._empty_tip)
        self._rebuild()

    # ---- 对外接口 -----------------------------------------------------

    def paths(self) -> list[str]:
        """返回当前已选择的文件路径列表。"""

        return list(self._paths)

    def set_paths(self, paths: list[str]) -> None:
        """整体替换文件列表（去重并保持顺序）。"""

        self._paths = list(dict.fromkeys(paths))
        self._rebuild()
        self.filesChanged.emit(self.paths())

    def add_paths(self, paths: list[str]) -> None:
        """追加文件；重复路径自动忽略。"""

        added = False
        for path in paths:
            if path and path not in self._paths:
                self._paths.append(path)
                added = True
        if added:
            self._rebuild()
            self.filesChanged.emit(self.paths())

    def remove_path(self, path: str) -> None:
        """移除单个文件。

        改造前选择器只能整体清空（选错一个就要全部重选），这里支持单文件移除。
        """

        if path in self._paths:
            self._paths.remove(path)
            self._meta.pop(path, None)
            self._rebuild()
            self.filesChanged.emit(self.paths())

    def set_meta(self, path: str, text: str) -> None:
        """设置某个文件的辅助信息，例如"4 个字段"。"""

        self._meta[path] = text
        self._rebuild()

    def clear(self) -> None:
        """清空全部选择。"""

        self._paths.clear()
        self._meta.clear()
        self._rebuild()
        self.filesChanged.emit([])

    # ---- 内部实现 -----------------------------------------------------

    def _open_dialog(self) -> None:
        """弹出多选文件对话框。"""

        files, _ = QFileDialog.getOpenFileNames(self, "选择表格文件", str(Path.home()), FILE_DIALOG_FILTER)
        if files:
            self.add_paths(files)

    def _rebuild(self) -> None:
        """重建文件行。

        先移除旧的 stretch 再插入新行，最后补回 stretch，
        保证文件行从上往下排列、不被拉伸变形。
        """

        while self._rows_layout.count() > 1:
            item = self._rows_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        for index, path in enumerate(self._paths):
            self._rows_layout.insertWidget(index, self._build_row(path))

        has_files = bool(self._paths)
        self._empty_tip.setVisible(not has_files)
        self._scroll.setVisible(has_files)

    def _build_row(self, path: str) -> QWidget:
        """构建单行文件卡片：文件名 + 元信息 + 移除按钮。"""

        row = QFrame()
        row.setObjectName("FileRow")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(SPACE["s"] + 2, 6, 6, 6)
        layout.setSpacing(SPACE["s"])

        name = QLabel(Path(path).name)
        name.setObjectName("FileName")
        name.setToolTip(path)
        name.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        layout.addWidget(name)

        meta = QLabel(self._meta.get(path, ""))
        meta.setObjectName("FileMeta")
        layout.addWidget(meta)

        remove = QPushButton("✕")
        remove.setObjectName("FileRemove")
        remove.setToolTip("移除该文件")
        remove.setCursor(Qt.CursorShape.PointingHandCursor)
        remove.setFixedSize(20, 20)
        # 用默认参数绑定当前 path，避免闭包捕获到循环变量。
        remove.clicked.connect(lambda _checked=False, target=path: self.remove_path(target))
        layout.addWidget(remove)
        return row

    # ---- 拖拽支持 -----------------------------------------------------

    def _dropped_paths(self, event) -> list[str]:
        """从拖拽事件里提取受支持的表格文件路径。"""

        if not event.mimeData().hasUrls():
            return []
        return collect_table_paths(event.mimeData().urls())

    def _rejected_paths(self, event) -> list[str]:
        """提取拖拽中存在的非法格式文件，使 dropEvent 能反馈原因。"""

        if not event.mimeData().hasUrls():
            return []
        return collect_rejected_table_paths(event.mimeData().urls())

    def dragEnterEvent(self, event) -> None:  # noqa: N802  (Qt 命名约定)
        if self._dropped_paths(event) or self._rejected_paths(event):
            event.acceptProposedAction()
            self.setProperty("dragActive", "true")
            refresh_style(self)

    def dragLeaveEvent(self, event) -> None:  # noqa: N802
        self.setProperty("dragActive", "false")
        refresh_style(self)

    def dropEvent(self, event) -> None:  # noqa: N802
        self.setProperty("dragActive", "false")
        refresh_style(self)
        paths = self._dropped_paths(event)
        rejected = self._rejected_paths(event)
        if paths:
            self.add_paths(paths)
        if rejected:
            self.filesRejected.emit(rejected)
        if paths or rejected:
            event.acceptProposedAction()


class SingleFilePicker(QFrame):
    """单文件选择卡（清洗 / 对比 / 图表页使用）。

    相比改造前的 `QLineEdit + 按钮`，这里把文件名、字段数、选择入口合并成一个控件，
    并同样支持拖拽，三个页面复用同一份实现。
    """

    fileChanged = Signal(str)
    # 单文件选择卡复用统一信号，页面负责记录日志并提示支持格式。
    filesRejected = Signal(list)

    def __init__(self, placeholder: str = "未选择文件", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)

        self._path = ""
        self._meta = ""

        layout = QHBoxLayout(self)
        layout.setContentsMargins(SPACE["m"] - 4, SPACE["s"] + 2, SPACE["m"] - 4, SPACE["s"] + 2)
        layout.setSpacing(SPACE["s"])

        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        self._name = QLabel(placeholder)
        self._name.setObjectName("FileName")
        self._meta_label = QLabel("可拖拽表格文件到此处")
        self._meta_label.setObjectName("FileMeta")
        text_box.addWidget(self._name)
        text_box.addWidget(self._meta_label)
        layout.addLayout(text_box, 1)

        button = QPushButton("选择文件")
        button.setObjectName("BtnSecondary")
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        button.clicked.connect(self._open_dialog)
        layout.addWidget(button)

    def path(self) -> str:
        """返回当前文件路径，未选择时返回空字符串。"""

        return self._path

    def set_path(self, path: str, emit: bool = True) -> None:
        """设置文件路径并刷新显示。

        参数 emit：是否发出 `fileChanged` 信号（页面回填时传 False 可避免重复解析）。
        """

        self._path = path
        self._meta = ""
        self._name.setText(Path(path).name if path else "未选择文件")
        self._meta_label.setText("可拖拽表格文件到此处" if not path else "解析中…")
        if emit:
            self.fileChanged.emit(path)

    def set_meta(self, text: str) -> None:
        """更新辅助信息，例如"N 个字段 · 公共字段 M 个"。"""

        self._meta = text
        if self._path:
            self._meta_label.setText(text)

    def set_error(self, message: str) -> None:
        """把辅助信息切换为错误提示（红色由 QSS 的 FileMeta 继承文字色控制）。"""

        self._meta_label.setText(message)
        self._meta_label.setStyleSheet(f"color:{COLOR['danger']};")

    def clear(self) -> None:
        """清空当前选择。"""

        self.set_path("")

    def _open_dialog(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "选择表格文件", str(Path.home()), FILE_DIALOG_FILTER)
        if file_path:
            self.set_path(file_path)

    def _dropped_paths(self, event) -> list[str]:
        """从拖拽事件里提取第一个受支持的表格文件。"""

        if not event.mimeData().hasUrls():
            return []
        return collect_table_paths(event.mimeData().urls())

    def _rejected_paths(self, event) -> list[str]:
        """提取拖拽中存在的非法格式文件，使 dropEvent 能反馈原因。"""

        if not event.mimeData().hasUrls():
            return []
        return collect_rejected_table_paths(event.mimeData().urls())

    def dragEnterEvent(self, event) -> None:  # noqa: N802  (Qt 命名约定)
        if self._dropped_paths(event) or self._rejected_paths(event):
            event.acceptProposedAction()
            self.setProperty("dragActive", "true")
            refresh_style(self)

    def dragLeaveEvent(self, event) -> None:  # noqa: N802
        self.setProperty("dragActive", "false")
        refresh_style(self)

    def dropEvent(self, event) -> None:  # noqa: N802
        self.setProperty("dragActive", "false")
        refresh_style(self)
        paths = self._dropped_paths(event)
        rejected = self._rejected_paths(event)
        if paths:
            # 单文件控件只接收第一个文件，多余文件忽略（页面按需提示）。
            self.set_path(paths[0])
        if rejected:
            self.filesRejected.emit(rejected)
        if paths or rejected:
            event.acceptProposedAction()


# ---------------------------------------------------------------------------
# 日志与反馈
# ---------------------------------------------------------------------------


class LogPanel(QFrame):
    """可折叠的操作日志面板。

    改造前日志是常驻的 QListWidget，占掉页面中部的黄金位置；
    这里默认折叠成一行摘要，点击标题才展开，把空间还给字段和配置。

    对外保留 `addItem()/clear()/scrollToBottom()`，主窗口的 `_append_log()`
    可以不改调用方式直接复用。
    """

    def __init__(self, title: str = "操作日志", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(SPACE["m"] - 4, SPACE["s"], SPACE["m"] - 4, SPACE["s"])
        layout.setSpacing(SPACE["s"] - 2)

        self._base_title = title

        header = QHBoxLayout()
        self._toggle = QPushButton()
        self._toggle.setObjectName("BtnGhost")
        self._toggle.setCursor(Qt.CursorShape.PointingHandCursor)
        self._toggle.clicked.connect(self._toggle_body)
        header.addWidget(self._toggle)
        header.addStretch(1)
        layout.addLayout(header)

        self._list = QListWidget()
        self._list.setObjectName("LogList")
        self._list.setMinimumHeight(96)
        self._list.setVisible(False)
        layout.addWidget(self._list)
        self._sync_toggle_text()

    def _sync_toggle_text(self) -> None:
        """同步折叠按钮上的箭头与条数统计。"""

        arrow = "▾" if self._list.isVisible() else "▸"
        self._toggle.setText(f"{arrow} {self._base_title} ({self._list.count()})")

    def _toggle_body(self) -> None:
        """展开 / 收起日志列表。"""

        visible = not self._list.isVisible()
        self._list.setVisible(visible)
        self._sync_toggle_text()
        if visible:
            self._list.scrollToBottom()

    def addItem(self, text: str) -> None:  # noqa: N802  (对齐 QListWidget 接口)
        """追加一条日志并刷新标题上的计数。"""

        self._list.addItem(text)
        self._sync_toggle_text()
        self._list.scrollToBottom()

    def clear(self) -> None:
        """清空日志。"""

        self._list.clear()
        self._sync_toggle_text()

    def scrollToBottom(self) -> None:  # noqa: N802
        """滚动到最新一行。"""

        self._list.scrollToBottom()


class FeedbackBar(QFrame):
    """页面底部反馈条：执行中 / 成功 / 失败 / 警告。

    改造前用 `QMessageBox` 弹模态框，一次任务连着几个校验会连环弹窗打断操作；
    这里改成常驻一条状态区，成功时提供"打开文件"快捷入口。
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("FeedbackBar")
        self.setProperty("state", "idle")

        self._action_slot: Callable[[], None] | None = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(SPACE["l"], SPACE["s"], SPACE["l"], SPACE["s"])
        layout.setSpacing(SPACE["s"])

        self._icon = QLabel("○")
        self._icon.setObjectName("FeedbackIcon")
        layout.addWidget(self._icon)

        self._text = QLabel("就绪")
        self._text.setObjectName("FeedbackText")
        layout.addWidget(self._text, 1)

        self._progress = QProgressBar()
        self._progress.setObjectName("FeedbackProgress")
        self._progress.setRange(0, 0)  # 0~0 表示不确定进度，Qt 会显示滚动动画
        self._progress.setTextVisible(False)
        self._progress.setFixedWidth(180)
        self._progress.setVisible(False)
        layout.addWidget(self._progress)

        self._action = QPushButton()
        self._action.setObjectName("BtnLink")
        self._action.setCursor(Qt.CursorShape.PointingHandCursor)
        self._action.setVisible(False)
        self._action.clicked.connect(self._run_action)
        layout.addWidget(self._action)

        self.show_idle()

    def _run_action(self) -> None:
        """执行当前绑定在"打开文件"按钮上的回调。"""

        if self._action_slot is not None:
            self._action_slot()

    def _set(
        self,
        state: str,
        icon: str,
        icon_color: str,
        message: str,
        busy: bool = False,
        action_text: str | None = None,
        action_slot: Callable[[], None] | None = None,
    ) -> None:
        """统一更新图标、文案、进度条和操作按钮。"""

        self.setProperty("state", state)
        self._icon.setText(icon)
        self._icon.setStyleSheet(f"color:{icon_color};")
        self._text.setText(message)
        self._text.setStyleSheet(f"color:{icon_color};")
        self._progress.setVisible(busy)

        self._action_slot = action_slot
        if action_text and action_slot:
            self._action.setText(action_text)
            self._action.setVisible(True)
        else:
            self._action.setVisible(False)
        # 动态属性变化必须手动刷新样式才会生效。
        refresh_style(self)

    def show_idle(self, message: str = "就绪") -> None:
        """空闲态：等待用户操作。"""

        self._set("idle", "○", COLOR["text_sub"], message)

    def show_busy(self, message: str) -> None:
        """执行中：显示不确定进度条。

        当前 Service 调用仍是主线程同步执行，进度条只表达"正在处理"，
        真正的异步与可取消在迭代计划 S2-01 中落地。
        """

        self._set("busy", "⟳", COLOR["primary"], message, busy=True)

    def show_success(self, message: str, action_text: str | None = None, action_slot: Callable[[], None] | None = None) -> None:
        """成功态：可选提供"打开文件"入口。"""

        self._set("success", "✓", COLOR["success"], message, action_text=action_text, action_slot=action_slot)

    def show_error(self, message: str) -> None:
        """失败态：红色提示，不再弹模态框。"""

        self._set("error", "✕", COLOR["danger"], message)

    def show_warning(self, message: str) -> None:
        """警告态：用于输入校验这类可恢复问题。"""

        self._set("warning", "!", COLOR["warning"], message)
