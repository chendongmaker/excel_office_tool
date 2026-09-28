"""界面主题模块：集中管理设计令牌与全局样式表。

改造前的问题：所有 QSS 都塞在 `MainWindow._apply_style()` 的一个字符串里，
颜色、间距、圆角散落各处，无法复用也无法统一维护。

改造后：
- `COLOR / SPACE / RADIUS / FONT` 是设计令牌，Python 代码里需要颜色时引用这里，
  不再写魔法值。
- `STYLE_SHEET` 是唯一的样式来源，覆盖导航、卡片、按钮、表格、输入控件、
  滚动条等所有控件类型。
- `apply_theme(app)` 是唯一入口，由 `MainWindow` 在初始化时调用。

层级约定：theme.py 只依赖 PySide6，不依赖任何业务模块，可被 UI 层任意引用。
"""

from __future__ import annotations

from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication, QWidget

# ---------------------------------------------------------------------------
# 设计令牌（Design Tokens）
# ---------------------------------------------------------------------------

# 颜色：与《前端界面优化方案》第二章规范一致，改一处即可全局生效。
COLOR: dict[str, str] = {
    "primary": "#2563EB",       # 主色：主按钮、选中态、链接
    "primary_hover": "#1D4ED8",  # 主色 hover
    "primary_soft": "#E8F0FF",   # 主色浅底：选中背景、徽章
    "success": "#16A34A",        # 成功
    "success_soft": "#EAF7EF",
    "danger": "#DC2626",         # 失败 / 危险操作
    "danger_soft": "#FDECEC",
    "warning": "#D97706",        # 警告
    "warning_soft": "#FEF6E7",
    "bg": "#F7F8FB",             # 窗口底色
    "bg_nav": "#FBFCFE",         # 导航底色
    "surface": "#FFFFFF",        # 卡片 / 输入框底色
    "border": "#E5E7EB",         # 常规描边
    "border_strong": "#D1D5DB",  # 输入框描边
    "text": "#1F2937",           # 正文
    "text_sub": "#64748B",       # 辅助文字
    "text_muted": "#9CA3AF",     # 更弱的提示（分组标题、占位符）
}

# 间距：统一内边距与外边距，避免每页各写一套数字。
SPACE: dict[str, int] = {"xs": 4, "s": 8, "m": 16, "l": 24, "xl": 32}

# 圆角：卡片用 m，输入控件用 s。
RADIUS: dict[str, int] = {"s": 6, "m": 10, "l": 14}

# 字体：Windows 中文环境优先微软雅黑，避免默认字体渲染发虚。
FONT_FAMILY = "Microsoft YaHei"
FONT_SIZE = 14


def build_style_sheet() -> str:
    """组装全局 QSS 字符串。

    使用 `.format(**COLOR)` 注入令牌，QSS 里不再出现硬编码颜色。
    Qt 的 QSS 不支持 box-shadow / transition，阴影通过控件层级和描边近似表达。
    """

    c = COLOR
    return f"""
    /* ---------- 基础 ----------
       注意：这里不给所有 QWidget 设背景色，否则卡片内部的 QLabel 会各自刷一层
       页面底色，把卡片切成一块块灰底。窗口底色只设在 QMainWindow / QStackedWidget 上，
       其余控件默认透明。 */
    QWidget {{
        color: {c['text']};
        font-size: {FONT_SIZE}px;
    }}
    QMainWindow, QStackedWidget {{ background: {c['bg']}; }}
    QMainWindow::separator {{ background: {c['border']}; }}
    QToolTip {{
        background: {c['text']};
        color: #FFFFFF;
        border: 0;
        padding: 6px 8px;
        border-radius: {RADIUS['s']}px;
    }}

    /* ---------- 菜单栏 ---------- */
    QMenuBar {{ background: {c['surface']}; border-bottom: 1px solid {c['border']}; padding: 2px 6px; }}
    QMenuBar::item {{ padding: 6px 12px; border-radius: {RADIUS['s']}px; }}
    QMenuBar::item:selected {{ background: {c['primary_soft']}; color: {c['primary']}; }}
    QMenu {{ background: {c['surface']}; border: 1px solid {c['border']}; padding: 6px; }}
    QMenu::item {{ padding: 7px 22px 7px 14px; border-radius: {RADIUS['s']}px; }}
    QMenu::item:selected {{ background: {c['primary_soft']}; color: {c['primary']}; }}

    /* ---------- 左侧导航 ---------- */
    QListWidget#Nav {{
        background: {c['bg_nav']};
        border: 0;
        border-right: 1px solid {c['border']};
        padding: 10px 8px;
        outline: 0;
    }}
    QListWidget#Nav::item {{
        padding: 10px 10px;
        border-radius: {RADIUS['s']}px;
        color: {c['text_sub']};
        margin: 1px 0;
    }}
    QListWidget#Nav::item:hover {{ background: #F1F5F9; color: {c['text']}; }}
    QListWidget#Nav::item:selected {{
        background: {c['primary_soft']};
        color: {c['primary']};
        font-weight: 600;
    }}
    /* 分组标题不可选中（flags 为 NoItemFlags），用 disabled 态弱化显示。 */
    QListWidget#Nav::item:disabled {{
        color: {c['text_muted']};
        font-size: 11px;
        background: transparent;
        padding-top: 14px;
        padding-bottom: 2px;
    }}

    /* ---------- 页头 ---------- */
    QWidget#PageHeader {{ background: {c['bg']}; }}
    QLabel#PageTitle {{ font-size: 20px; font-weight: 700; }}
    QLabel#PageDesc {{ color: {c['text_sub']}; font-size: 12.5px; }}
    QLabel#HeroTitle {{ font-size: 26px; font-weight: 700; }}
    QLabel#SectionLabel {{ font-weight: 600; font-size: 13.5px; }}
    QLabel#Muted {{ color: {c['text_sub']}; font-size: 12.5px; }}
    QLabel#MutedSmall {{ color: {c['text_muted']}; font-size: 11.5px; }}

    /* ---------- 卡片 ---------- */
    QFrame#Card {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: {RADIUS['m']}px;
    }}
    QLabel#CardTitle {{ font-size: 13.5px; font-weight: 600; }}
    QLabel#StepBadge {{
        background: {c['primary']};
        color: #FFFFFF;
        font-size: 11px;
        font-weight: 700;
        border-radius: 9px;
        min-width: 18px;
        max-width: 18px;
        min-height: 18px;
        max-height: 18px;
        qproperty-alignment: AlignCenter;
    }}
    QLabel#ChipInfo {{
        background: {c['primary_soft']};
        color: {c['primary']};
        border-radius: 9px;
        padding: 2px 10px;
        font-size: 11.5px;
        font-weight: 600;
    }}
    QLabel#StatValue {{ font-size: 22px; font-weight: 700; }}
    QLabel#StatKey {{ color: {c['text_sub']}; font-size: 12px; }}

    /* 首页功能入口卡：用 QPushButton 承载卡片，hover 时换蓝色描边表达"可点击"。 */
    QPushButton#FunctionCard {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: {RADIUS['m']}px;
        text-align: left;
        padding: 0;
    }}
    QPushButton#FunctionCard:hover {{ border: 1px solid {c['primary']}; background: #FCFDFF; }}
    QLabel#FunctionTitle {{ font-size: 14px; font-weight: 600; }}

    /* ---------- 按钮三级层级 ----------
       注意：这里不写 min-height/max-height。Qt 的 QSS 尺寸属性会覆盖
       setMinimumHeight()，曾导致首页功能卡被压回默认按钮高度。 */
    QPushButton {{
        background: {c['surface']};
        color: {c['text']};
        border: 1px solid {c['border_strong']};
        border-radius: {RADIUS['s']}px;
        padding: 8px 16px;
    }}
    QPushButton:hover {{ border-color: {c['primary']}; color: {c['primary']}; }}
    QPushButton:disabled {{ background: #F1F5F9; color: {c['text_muted']}; border-color: {c['border']}; }}

    /* Primary：每页唯一的主操作按钮，蓝底白字且不铺满整行。 */
    QPushButton#BtnPrimary {{
        background: {c['primary']};
        color: #FFFFFF;
        border: 0;
        font-weight: 600;
        padding: 9px 20px;
    }}
    QPushButton#BtnPrimary:hover {{ background: {c['primary_hover']}; color: #FFFFFF; }}
    QPushButton#BtnPrimary:disabled {{ background: #BFD3F8; color: #FFFFFF; }}

    /* Secondary：白底蓝边，用于"选择文件"这类次要动作。 */
    QPushButton#BtnSecondary {{
        background: {c['surface']};
        color: {c['primary']};
        border: 1px solid {c['primary']};
        font-weight: 600;
    }}
    QPushButton#BtnSecondary:hover {{ background: {c['primary_soft']}; }}

    /* Ghost：无边框文字按钮，用于全选/反选/展开日志等低权重动作。 */
    QPushButton#BtnGhost {{
        background: transparent;
        border: 0;
        color: {c['text_sub']};
        padding: 4px 8px;
    }}
    QPushButton#BtnGhost:hover {{ color: {c['primary']}; background: {c['primary_soft']}; }}

    /* Danger：删除类动作，二次确认后才会真正执行。 */
    QPushButton#BtnDanger {{ background: transparent; border: 0; color: {c['danger']}; padding: 4px 8px; }}
    QPushButton#BtnDanger:hover {{ background: {c['danger_soft']}; }}

    QPushButton#BtnLink {{ background: transparent; border: 0; color: {c['primary']}; padding: 2px 6px; }}
    QPushButton#BtnLink:hover {{ text-decoration: underline; }}

    /* 清洗规则开关：整行可点的可勾选按钮，勾选=启用（浅蓝底），未勾选=禁用（灰色删除线）。 */
    QPushButton#RuleCheck {{
        background: transparent;
        border: 1px solid transparent;
        border-radius: {RADIUS['s']}px;
        text-align: left;
        padding: 6px 10px;
        color: {c['text_sub']};
    }}
    QPushButton#RuleCheck:hover {{ background: #F1F5F9; }}
    QPushButton#RuleCheck:checked {{
        background: {c['primary_soft']};
        border: 1px solid #C7D2FE;
        color: {c['primary']};
        font-weight: 600;
    }}
    QPushButton#RuleCheck:!checked {{ color: {c['text_muted']}; text-decoration: line-through; }}

    QPushButton#FileRemove {{ background: transparent; border: 0; color: {c['text_muted']}; font-size: 14px; padding: 0 4px; }}
    QPushButton#FileRemove:hover {{ color: {c['danger']}; }}

    /* ---------- 输入控件 ---------- */
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
        background: {c['surface']};
        border: 1px solid {c['border_strong']};
        border-radius: {RADIUS['s']}px;
        padding: 7px 10px;
        selection-background-color: {c['primary_soft']};
        selection-color: {c['primary']};
    }}
    QLineEdit:focus, QComboBox:focus {{ border: 1px solid {c['primary']}; }}
    QLineEdit[readOnly="true"] {{ background: #F8FAFC; color: {c['text_sub']}; }}
    QLineEdit:disabled, QComboBox:disabled {{ background: #F1F5F9; color: {c['text_muted']}; }}
    /* 校验失败时给输入框加 error 动态属性，描边变红。 */
    QLineEdit[error="true"] {{ border: 1px solid {c['danger']}; background: {c['danger_soft']}; }}

    QComboBox::drop-down {{ border: 0; width: 22px; }}
    QComboBox QAbstractItemView {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        selection-background-color: {c['primary_soft']};
        selection-color: {c['primary']};
        outline: 0;
    }}

    /* ---------- 勾选与单选 ---------- */
    QCheckBox, QRadioButton {{ spacing: 7px; padding: 2px 0; }}
    QCheckBox::indicator, QRadioButton::indicator {{ width: 15px; height: 15px; }}
    QCheckBox::indicator {{ border: 1.5px solid {c['border_strong']}; border-radius: 4px; background: {c['surface']}; }}
    QCheckBox::indicator:hover {{ border-color: {c['primary']}; }}
    QCheckBox::indicator:checked {{ background: {c['primary']}; border-color: {c['primary']}; }}
    QRadioButton::indicator {{ border: 1.5px solid {c['border_strong']}; border-radius: 8px; background: {c['surface']}; }}
    QRadioButton::indicator:hover {{ border-color: {c['primary']}; }}
    QRadioButton::indicator:checked {{ border: 4px solid {c['primary']}; background: {c['surface']}; }}

    /* ---------- 表格 ---------- */
    QTableWidget, QTableView {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: {RADIUS['s']}px;
        gridline-color: #F1F5F9;
        outline: 0;
    }}
    QTableWidget::item {{ padding: 6px; }}
    QTableWidget::item:selected {{ background: {c['primary_soft']}; color: {c['primary']}; }}
    QHeaderView::section {{
        background: #F8FAFC;
        color: {c['text_sub']};
        border: 0;
        border-bottom: 1px solid {c['border']};
        border-right: 1px solid {c['border']};
        padding: 8px 10px;
        font-weight: 600;
    }}
    QTableCornerButton::section {{ background: #F8FAFC; border: 0; }}

    /* ---------- 列表（日志 / 最近任务） ---------- */
    QListWidget#LogList, QListWidget#RecentList {{
        background: #FAFBFC;
        border: 1px solid {c['border']};
        border-radius: {RADIUS['s']}px;
        padding: 6px;
        outline: 0;
    }}
    QListWidget#LogList::item {{ padding: 4px 6px; color: {c['text_sub']}; }}
    QListWidget#RecentList::item {{ padding: 8px 6px; }}

    /* ---------- 拖拽与空状态 ---------- */
    QFrame#DropZone {{
        background: #FAFBFF;
        border: 1.5px dashed #C7D2FE;
        border-radius: {RADIUS['m']}px;
    }}
    QFrame#DropZone[dragActive="true"] {{ background: {c['primary_soft']}; border-color: {c['primary']}; }}
    QFrame#EmptyState {{ background: #FCFCFD; border: 1px dashed {c['border']}; border-radius: {RADIUS['m']}px; }}
    QLabel#EmptyTitle {{ color: {c['text_sub']}; font-size: 13.5px; }}
    QLabel#EmptyHint {{ color: {c['text_muted']}; font-size: 11.5px; }}

    /* 文件小卡片：一行一个文件，带删除按钮。 */
    QFrame#FileRow {{ background: #FAFBFC; border: 1px solid {c['border']}; border-radius: {RADIUS['s']}px; }}
    QLabel#FileName {{ font-weight: 600; }}
    QLabel#FileMeta {{ color: {c['text_sub']}; font-size: 11.5px; }}

    /* ---------- 反馈条 ---------- */
    QFrame#FeedbackBar {{
        background: {c['surface']};
        border-top: 1px solid {c['border']};
    }}
    QFrame#FeedbackBar[state="success"] {{ background: {c['success_soft']}; border-top: 1px solid #CBEBD6; }}
    QFrame#FeedbackBar[state="error"] {{ background: {c['danger_soft']}; border-top: 1px solid #F5D2D2; }}
    QFrame#FeedbackBar[state="warning"] {{ background: {c['warning_soft']}; border-top: 1px solid #F5E3C3; }}
    QLabel#FeedbackIcon {{ font-size: 14px; font-weight: 700; }}
    QLabel#FeedbackText {{ font-size: 12.5px; }}
    QProgressBar#FeedbackProgress {{
        background: #E2E8F0;
        border: 0;
        border-radius: 3px;
        height: 6px;
        max-height: 6px;
    }}
    QProgressBar#FeedbackProgress::chunk {{ background: {c['primary']}; border-radius: 3px; }}

    /* ---------- 徽章 ---------- */
    QLabel#BadgeSuccess {{ background: {c['success_soft']}; color: {c['success']}; border-radius: 9px; padding: 1px 9px; font-size: 11.5px; font-weight: 600; }}
    QLabel#BadgeFail {{ background: {c['danger_soft']}; color: {c['danger']}; border-radius: 9px; padding: 1px 9px; font-size: 11.5px; font-weight: 600; }}
    QLabel#BadgePartial {{ background: {c['warning_soft']}; color: {c['warning']}; border-radius: 9px; padding: 1px 9px; font-size: 11.5px; font-weight: 600; }}

    /* ---------- 滚动条：8px 细条，避免系统原生粗滚动条破坏整体感 ---------- */
    QScrollBar:vertical {{ background: transparent; width: 8px; margin: 2px; }}
    QScrollBar::handle:vertical {{ background: #CBD5E1; border-radius: 4px; min-height: 28px; }}
    QScrollBar::handle:vertical:hover {{ background: #94A3B8; }}
    QScrollBar:horizontal {{ background: transparent; height: 8px; margin: 2px; }}
    QScrollBar::handle:horizontal {{ background: #CBD5E1; border-radius: 4px; min-width: 28px; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; width: 0; }}
    QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

    /* ---------- 分隔与滚动区 ---------- */
    QSplitter::handle {{ background: {c['border']}; width: 1px; }}
    QScrollArea {{ border: 0; background: transparent; }}
    QScrollArea > QWidget > QWidget {{ background: transparent; }}
    """


def apply_theme(widget: QWidget | QApplication) -> None:
    """把主题应用到整个应用（QApplication）或单个窗口（QWidget）。

    统一设置字体与样式表，保证所有页面控件外观一致。
    """

    font = QFont(FONT_FAMILY, 10)
    widget.setFont(font)
    widget.setStyleSheet(build_style_sheet())


def refresh_style(widget: QWidget) -> None:
    """动态属性变更后刷新控件样式。

    Qt 的 QSS 属性选择器（如 `[state="success"]`）不会自动重绘，
    必须 unpolish + polish 一次才会重新计算样式。
    """

    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()
