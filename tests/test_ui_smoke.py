"""UI 层冒烟测试。

目的：界面改造涉及所有页面的控件结构，单元测试只覆盖 core 层，
需要一组"能真正把主窗口建起来 + 关键交互链路能跑通"的测试兜底，
避免以后再改界面时出现"导入报错/回调名写错"这类低级回归。

做法：
- 用 offscreen 平台离屏创建窗口，不依赖桌面显示环境，可在 CI 里跑。
- 断言只针对结构和状态（页面数量、字段解析结果、下选项、反馈条状态），
  不断言像素渲染，避免受字体/主题差异影响。
"""

from __future__ import annotations

import os

import pandas as pd
import pytest

# Qt 平台必须在创建 QApplication 之前指定，离屏平台无需显示器。
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(scope="module")
def qt_app():
    """创建（或复用）QApplication 实例。"""

    from PySide6.QtWidgets import QApplication

    instance = QApplication.instance() or QApplication([])
    yield instance


@pytest.fixture()
def window(qt_app, tmp_path):
    """创建主窗口，并把历史记录重定向到临时目录，避免污染用户真实历史。"""

    from excel_helper.storage.history import HistoryStore
    from excel_helper.ui.main_window import MainWindow

    main_window = MainWindow()
    main_window.history = HistoryStore(tmp_path / "history.json")
    main_window._refresh_history()
    yield main_window
    main_window.close()


@pytest.fixture()
def sample_files(tmp_path) -> list[str]:
    """构造两个字段部分重叠的示例表格，用于验证字段解析与合并配置。"""

    first = tmp_path / "1月销售.xlsx"
    second = tmp_path / "2月销售.xlsx"
    pd.DataFrame({"商品ID": ["A001", "A002"], "商品名称": ["苹果", "香蕉"], "库存": [120, 80]}).to_excel(
        first, index=False
    )
    pd.DataFrame({"商品ID": ["A001", "A002"], "商品名称": ["苹果", "香蕉"], "销量": [30, 25]}).to_excel(
        second, index=False
    )
    return [str(first), str(second)]


def test_main_window_builds_all_pages(window):
    """主窗口应包含 6 个页面，导航映射指向正确的页面索引。"""

    assert window.stack.count() == 6
    # 导航包含 2 个不可点击的分组标题，因此行数多于页面数。
    assert window.nav.count() == 8
    assert sorted(window._nav_to_stack.values()) == [0, 1, 2, 3, 4, 5]
    # 初始停留在首页。
    assert window.stack.currentIndex() == 0


def test_merge_page_parses_fields_and_common_keys(window, sample_files):
    """合并页选择文件后：字段表填充、公共字段下拉可用、主按钮解禁。"""

    window.merge_files.set_paths(sample_files)

    # 字段解析结果：3 个公共字段 + 2 个部分文件字段。
    field_names = {field.name for field in window.merge_field_selector.fields}
    assert field_names == {"商品ID", "商品名称", "库存", "销量"}
    assert window.merge_field_selector.has_fields() is True

    # 公共字段默认勾选，部分文件字段默认不勾选。
    selected = window.merge_field_selector.selected_fields()
    assert set(selected) == {"商品ID", "商品名称"}

    # 匹配字段下拉只列公共字段，且默认未启用（当前是"按行纵向合并"）。
    key_options = [window.merge_key_input.itemText(i) for i in range(window.merge_key_input.count())]
    assert set(key_options) == {"商品ID", "商品名称"}
    assert window.merge_key_input.isEnabled() is False
    window.merge_key_radio.setChecked(True)
    assert window.merge_key_input.isEnabled() is True

    # 主按钮在选择文件后才可用。
    assert window.merge_run_button.isEnabled() is True
    assert window.merge_file_chip.text() == "共 2 个文件"


def test_merge_file_removal_keeps_other_files(window, sample_files):
    """单个文件可移除后，剩余文件仍保留（改造前只能整体清空）。"""

    window.merge_files.set_paths(sample_files)
    window.merge_files.remove_path(sample_files[0])

    assert window.merge_files.paths() == [sample_files[1]]
    assert window.merge_file_chip.text() == "共 1 个文件"


def test_chart_page_fills_field_combos(window, sample_files):
    """图表页选择文件后，X/Y/分组下拉应填入真实表头，不再需要手输。"""

    window.chart_picker.set_path(sample_files[0])

    expected = ["商品ID", "商品名称", "库存"]
    assert [window.chart_x_input.itemText(i) for i in range(window.chart_x_input.count())] == expected
    assert [window.chart_y_input.itemText(i) for i in range(window.chart_y_input.count())] == expected
    # 分组下拉第一项代表"不分组"，data 为 None。
    assert window.chart_group_input.itemData(0) is None
    assert window.chart_group_input.count() == len(expected) + 1
    assert window.chart_run_button.isEnabled() is True


def test_chart_page_prefers_numeric_y_axis(window, sample_files):
    """图表 Y 轴必须能求和：默认值应落在数值列，且文本列会被提前拦截。

    背景：core 的 chart_factory 使用 groupby().sum()，文本列会在执行阶段抛
    "Cannot use numeric_only=True ..." 的 pandas 报错，所以 UI 层要提前把关。
    """

    window.chart_picker.set_path(sample_files[0])

    assert window._is_numeric_dtype(window._chart_column_dtype("库存")) is True
    assert window._is_numeric_dtype(window._chart_column_dtype("商品名称")) is False
    # 示例文件里唯一的数值列是"库存"，应该被自动选中。
    assert window.chart_y_input.currentText() == "库存"

    # 手动改成文本列后执行，应给出可读的警告而不是异常堆栈。
    window.chart_y_input.setCurrentText("商品名称")
    window._run_chart()

    assert window.feedback.property("state") == "warning"


def test_compare_page_requires_two_files(window, sample_files):
    """对比页只有在两个文件都选好后才允许执行。"""

    window.compare_left_picker.set_path(sample_files[0])
    assert window.compare_run_button.isEnabled() is False

    window.compare_right_picker.set_path(sample_files[1])
    assert window.compare_run_button.isEnabled() is True
    key_options = [window.compare_key_input.itemText(i) for i in range(window.compare_key_input.count())]
    assert set(key_options) == {"商品ID", "商品名称"}


def test_warning_goes_to_feedback_bar(window):
    """校验失败走反馈条（警告态），不再弹模态框阻塞界面。"""

    from PySide6.QtWidgets import QLabel

    window._warn("请先选择要合并的文件")

    assert window.feedback.property("state") == "warning"
    message = window.feedback.findChild(QLabel, "FeedbackText")
    assert message is not None and "请先选择要合并的文件" in message.text()


def test_clean_rules_can_be_disabled(window, sample_files):
    """清洗规则开关可关闭，且默认全部启用（保持与改造前行为一致）。"""

    assert all(toggle.isChecked() for toggle in window.clean_rule_checks.values())

    window.clean_rule_checks["drop_duplicates"].setChecked(False)

    assert window.clean_rule_checks["drop_duplicates"].isChecked() is False
    assert window.clean_rule_checks["strip_text"].isChecked() is True


def test_history_table_renders_records_and_empty_state(window):
    """历史记录写入后表格可渲染，清空后回到空状态。"""

    from excel_helper.models.task import TaskRecord, TaskStatus

    window.history.add(
        TaskRecord(
            name="纵向合并",
            status=TaskStatus.SUCCESS,
            outputs=[window.history.path.parent / "合并.xlsx"],
            message="已导出",
        )
    )
    window._refresh_history()

    assert window.history_table.rowCount() == 1
    assert window.history_table.item(0, 1).text() == "纵向合并"
    assert window.history_table.item(0, 2).text() == "成功"
    # 有记录时隐藏空状态（窗口未 show，只能用 isHidden 判断显隐意图）。
    assert window.history_empty.isHidden() is True

    window.history.clear()
    window._refresh_history()

    assert window.history_table.rowCount() == 0
    assert window.home_history.rowCount() == 0
    assert window.history_empty.isHidden() is False
