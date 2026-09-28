from __future__ import annotations

import pandas as pd

from excel_helper.chart.chart_factory import save_chart
from excel_helper.core.excel_clean import apply_clean_rules
from excel_helper.core.excel_compare import compare_by_key
from excel_helper.core.field_parser import build_field_coverage, parse_files_metadata
from excel_helper.core.excel_merge import merge_by_key, merge_by_rows
from excel_helper.models.rule import CleanAction, CleanRule


def test_merge_by_rows_adds_source_file(tmp_path):
    """纵向合并时应追加“来源文件”，让用户能追溯每行来自哪个 Excel。"""

    first = tmp_path / "jan.xlsx"
    second = tmp_path / "feb.xlsx"
    pd.DataFrame({"商品": ["A"], "销量": [100]}).to_excel(first, index=False)
    pd.DataFrame({"商品": ["B"], "销量": [200]}).to_excel(second, index=False)

    result = merge_by_rows([first, second])

    assert list(result["商品"]) == ["A", "B"]
    assert list(result["来源文件"]) == ["jan.xlsx", "feb.xlsx"]


def test_parse_field_coverage_marks_common_and_partial_fields(tmp_path):
    """字段解析应区分公共字段和部分文件字段，并设置默认勾选状态。"""

    first = tmp_path / "a.xlsx"
    second = tmp_path / "b.xlsx"
    pd.DataFrame({"商品ID": [1], "商品名称": ["A"], "成本价": [10]}).to_excel(first, index=False)
    pd.DataFrame({"商品ID": [2], "商品名称": ["B"], "利润": [5]}).to_excel(second, index=False)

    fields = build_field_coverage(parse_files_metadata([first, second]))
    by_name = {field.name: field for field in fields}

    assert by_name["商品ID"].files == 2
    assert by_name["商品ID"].selected is True
    assert by_name["成本价"].files == 1
    assert by_name["成本价"].selected is False


def test_merge_by_key(tmp_path):
    """按主键匹配合并时，应把不同文件中的字段横向拼接到同一行。"""

    products = tmp_path / "products.xlsx"
    sales = tmp_path / "sales.xlsx"
    pd.DataFrame({"商品ID": [1001, 1002], "商品名": ["手机", "电脑"]}).to_excel(products, index=False)
    pd.DataFrame({"商品ID": [1001, 1002], "销量": [300, 500]}).to_excel(sales, index=False)

    result = merge_by_key([products, sales], key="商品ID")

    assert list(result.columns) == ["商品ID", "商品名", "销量"]
    assert result.loc[result["商品ID"] == 1002, "销量"].item() == 500


def test_merge_by_rows_uses_selected_columns(tmp_path):
    """用户只勾选部分字段时，合并结果不应导出未勾选列。"""

    first = tmp_path / "jan.xlsx"
    second = tmp_path / "feb.xlsx"
    pd.DataFrame({"商品": ["A"], "销量": [100], "内部备注": ["x"]}).to_excel(first, index=False)
    pd.DataFrame({"商品": ["B"], "销量": [200], "利润": [20]}).to_excel(second, index=False)

    result = merge_by_rows([first, second], columns=["商品", "销量"])

    assert list(result.columns) == ["来源文件", "商品", "销量"]


def test_clean_rules():
    """清洗规则链应能完成去重、去空格、改列名、数字和日期标准化。"""

    frame = pd.DataFrame(
        {
            " 商品 ": ["  苹果手机  ", "  苹果手机  ", None],
            "金额": ["￥1,200元", "￥1,200元", None],
            "日期": ["2026/08/01", "2026/08/01", None],
        }
    )

    result = apply_clean_rules(
        frame,
        [
            CleanRule(CleanAction.DROP_EMPTY_ROWS),
            CleanRule(CleanAction.DROP_DUPLICATES),
            CleanRule(CleanAction.STRIP_TEXT),
            CleanRule(CleanAction.RENAME_COLUMNS, rename_map={" 商品 ": "商品"}),
            CleanRule(CleanAction.NORMALIZE_NUMBERS, columns=["金额"]),
            CleanRule(CleanAction.NORMALIZE_DATES, columns=["日期"]),
        ],
    )

    assert len(result) == 1
    assert result["商品"].item() == "苹果手机"
    assert result["金额"].item() == 1200
    assert result["日期"].item() == "2026-08-01"


def test_clean_rules_survive_trimmed_column_names():
    """列名去空格后，后续字段级规则仍能匹配用户原先选择的字段。"""

    frame = pd.DataFrame({" 商品 ": ["  A  "], "金额": [None]})

    result = apply_clean_rules(
        frame,
        [
            CleanRule(CleanAction.STRIP_TEXT, columns=[" 商品 "]),
            CleanRule(CleanAction.FILL_EMPTY, columns=[" 商品 "], fill_value="未知"),
        ],
    )

    assert list(result.columns) == ["商品", "金额"]
    assert result["商品"].item() == "A"


def test_compare_by_key(tmp_path):
    """按主键对比应输出新增、删除和修改三类结果。"""

    left = tmp_path / "old.xlsx"
    right = tmp_path / "new.xlsx"
    pd.DataFrame({"商品ID": [1, 2], "销量": [100, 200]}).to_excel(left, index=False)
    pd.DataFrame({"商品ID": [1, 3], "销量": [150, 300]}).to_excel(right, index=False)

    result = compare_by_key(left, right, "商品ID")

    assert list(result["added"]["商品ID"]) == [3]
    assert list(result["removed"]["商品ID"]) == [2]
    assert result["changed"].iloc[0]["变化"] == 50


def test_compare_by_key_uses_selected_compare_columns(tmp_path):
    """对比字段来自字段选择器，未勾选字段即使变化也不应出现在修改表中。"""

    left = tmp_path / "old.xlsx"
    right = tmp_path / "new.xlsx"
    pd.DataFrame({"商品ID": [1], "销量": [100], "备注": ["旧"]}).to_excel(left, index=False)
    pd.DataFrame({"商品ID": [1], "销量": [150], "备注": ["新"]}).to_excel(right, index=False)

    result = compare_by_key(left, right, "商品ID", compare_columns=["销量"])

    assert list(result["changed"]["字段"]) == ["销量"]


def test_save_chart(tmp_path):
    """图表生成应输出非空 PNG 文件。"""

    output = tmp_path / "chart.png"
    frame = pd.DataFrame({"月份": ["1月", "2月"], "销售额": [100, 180]})

    result = save_chart(frame, "bar", "月份", "销售额", output)

    assert result.exists()
    assert result.stat().st_size > 0
