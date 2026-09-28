from __future__ import annotations

from pathlib import Path

import matplotlib

# 使用 Agg 后端可以在没有桌面显示器的测试/打包环境中生成 PNG。
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd


def save_chart(
    frame: pd.DataFrame,
    chart_type: str,
    x_column: str,
    y_column: str,
    output_path: str | Path,
    group_column: str | None = None,
    title: str = "",
) -> Path:
    """根据用户选择的字段生成图表 PNG。

    上游：
    - `ChartService.create_chart()` 传入 DataFrame 和 UI 表单参数。

    参数：
    - chart_type：bar、line、pie。
    - x_column：维度字段，例如“月份”或“商品分类”。
    - y_column：指标字段，例如“销售额”。
    - group_column：可选分组字段，例如“渠道”，用于生成多系列柱状/折线图。
    """

    if x_column not in frame.columns or y_column not in frame.columns:
        raise ValueError("X轴或Y轴字段不存在")

    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    # 中文字体优先用 Windows 常见字体，缺失时回退到 matplotlib 自带字体。
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(10, 5.8), dpi=140)
    chart_type = chart_type.lower()

    if chart_type == "pie":
        # 饼图只需要一个维度和一个指标，先按维度聚合再计算占比。
        data = frame.groupby(x_column, dropna=False)[y_column].sum(numeric_only=True)
        ax.pie(data.values, labels=data.index.astype(str), autopct="%1.1f%%", startangle=90)
        ax.axis("equal")
    elif group_column and group_column in frame.columns:
        # 有分组字段时转成透视表，每个分组是一条线或一组柱。
        pivot = frame.pivot_table(index=x_column, columns=group_column, values=y_column, aggfunc="sum")
        if chart_type == "line":
            pivot.plot(ax=ax, marker="o")
        else:
            pivot.plot(ax=ax, kind="bar")
    else:
        # 无分组字段时，按 X 轴字段聚合成单系列图表。
        data = frame.groupby(x_column, dropna=False)[y_column].sum(numeric_only=True)
        if chart_type == "line":
            ax.plot(data.index.astype(str), data.values, marker="o", linewidth=2.4)
        else:
            ax.bar(data.index.astype(str), data.values)

    ax.set_title(title or f"{y_column} 按 {x_column} 统计")
    if chart_type != "pie":
        ax.set_xlabel(x_column)
        ax.set_ylabel(y_column)
        ax.grid(axis="y", linestyle="--", alpha=0.35)
        ax.tick_params(axis="x", rotation=25)
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    # 主动关闭 figure，避免批量生成图表时占用内存。
    plt.close(fig)
    return path
