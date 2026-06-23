"""
reli-stat — 可靠性统计 Excel 工具

单文件实现。核心功能:
  - 读取 Excel 测试数据
  - 按 group 分组计算 CDF (Median Rank) 和 Weibull 变换
  - 输出带统计汇总 sheet + 每数据列独立 sheet + 散点图的 Excel
  - 独立绘图函数 add_excel_chart() 可直接在任何地方复用

依赖: openpyxl, pandas
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    NamedStyle,
    PatternFill,
    Side,
    numbers,
)
from openpyxl.utils import get_column_letter
from openpyxl.chart.layout import Layout, ManualLayout
# ──────────────────────────────────────────────
# 常量
# ──────────────────────────────────────────────

# 10 色调色板（循环使用）
COLOR_PALETTE = [
    "4472C4",  # 蓝
    "ED7D31",  # 橙
    "70AD47",  # 绿
    "FFC000",  # 金
    "5B9BD5",  # 浅蓝
    "A5A5A5",  # 灰
    "264478",  # 深蓝
    "9B59B6",  # 紫
    "E74C3C",  # 红
    "1ABC9C",  # 青
]

# 10 种 marker 形状（循环使用）
MARKER_SYMBOLS = [
    "circle",
    "square",
    "diamond",
    "triangle",
    "plus",
    "star",
    "x",
    "dash",
    "dot",
    "auto",
]

# 汇总统计表头
SUMMARY_HEADERS = [
    "测试项",
    "总模块数",
    "Group",
    "均值",
    "标准差",
    "25%分位",
    "75%分位",
    "中位数",
    "最小值",
    "最大值",
    "变异系数(CV%)",
    "Weibull β",
    "Weibull η",
    "拟合 R²",
    "Limit",
    "limit处 CDF",
]

# Header style
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


# ──────────────────────────────────────────────
# 统计计算（纯函数，零 IO）
# ──────────────────────────────────────────────


def empirical_cdf(values: pd.Series) -> pd.Series:
    """
    计算经验 CDF = rank / n。

    rank 使用 average method（相同值取平均秩次）。
    返回 Series，索引同输入。
    """
    n = len(values)
    if n == 0:
        return pd.Series(dtype=float)

    ranks = values.rank(method="average")
    return ranks / n


def median_rank(values: pd.Series) -> pd.Series:
    """
    计算 Median Rank (Benard's approximation)。

    MR_i = (rank_i - 0.3) / (n + 0.4)

    用于 Weibull 变换。返回 Series，索引同输入。
    """
    n = len(values)
    if n == 0:
        return pd.Series(dtype=float)

    ranks = values.rank(method="average")
    mr = (ranks - 0.3) / (n + 0.4)
    return mr


def _add_gridlines(chart: ScatterChart) -> None:
    """给散点图添加主网格线和次网格线（浅灰）。"""
    from openpyxl.chart.axis import ChartLines
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties

    for axis in (chart.x_axis, chart.y_axis):
        # 主网格线
        axis.majorGridlines = ChartLines()
        sp_pr = GraphicalProperties()
        sp_pr.ln = LineProperties(solidFill="D9D9D9", w=6350)
        axis.majorGridlines.spPr = sp_pr
        # 次网格线（更淡）
        axis.minorGridlines = ChartLines()
        sp_pr2 = GraphicalProperties()
        sp_pr2.ln = LineProperties(solidFill="ECECEC", w=3175)
        axis.minorGridlines.spPr = sp_pr2


def weibull_transform(cdf: pd.Series) -> pd.Series:
    """
    Weibull 变换: W = ln(-ln(1 - CDF))

    CDF = 0 或 1 时结果为 ±inf → 保留为 NaN。
    """
    import numpy as np

    # 边界保护: 0 < CDF < 1
    valid = (cdf > 0) & (cdf < 1)
    result = pd.Series(np.nan, index=cdf.index, dtype=float)
    result[valid] = np.log(-np.log(1 - cdf[valid]))
    return result


def fit_weibull(data: pd.Series, mr: pd.Series) -> tuple[float, float, float] | None:
    """
    Weibull 分布拟合。

    对 ln(data) 和 weibull(mr) 做线性回归：
        ln(-ln(1-MR)) = beta * ln(x) - beta * ln(eta)

    返回 (beta, eta, R²)。data 含非正值或样本不足时返回 None。
    """
    import numpy as np

    valid = (data > 0) & (mr > 0) & (mr < 1)
    x = data[valid]
    y = mr[valid]

    if len(x) < 3:
        return None

    ln_x = np.log(x.values)
    ln_weibull = np.log(-np.log(1 - y.values))

    A = np.vstack([ln_x, np.ones(len(ln_x))]).T
    slope, intercept = np.linalg.lstsq(A, ln_weibull, rcond=None)[0]

    beta = slope
    eta = np.exp(-intercept / beta) if beta != 0 else float("nan")

    y_pred = slope * ln_x + intercept
    ss_res = np.sum((ln_weibull - y_pred) ** 2)
    ss_tot = np.sum((ln_weibull - np.mean(ln_weibull)) ** 2)
    r_squared = 1 - ss_res / ss_tot if ss_tot != 0 else float("nan")

    return (round(beta, 4), round(eta, 4), round(r_squared, 4))


def get_limit(
    variable: str, limit_map: dict[str, float] | None
) -> float | None:
    """
    用正则匹配测试项名称查 limit_map。

    limit_map: {r"Vth.*": 3.0, r"BVdss": 650, ...}
    key 是正则表达式，匹配 variable 名，首个命中返回对应值；未命中返回 None。
    """
    if not limit_map:
        return None
    for pattern_str, value in limit_map.items():
        if re.search(pattern_str, variable):
            return float(value)
    return None


def compute_statistics(
    df: pd.DataFrame,
    id_col: str,
    group_col: str,
    data_cols: list[str],
    limit_map: dict[str, float] | None = None,
) -> pd.DataFrame:
    """
    计算所有数据列的 CDF、Weibull 变换 和 limit。

    对每个 data_col，按 group 分组计算 median rank CDF。

    返回长表格式 DataFrame，含列:
        id, group, variable, data, CDF, weibull, limit
    """
    frames = []

    for col in data_cols:
        sub = df[[id_col, group_col, col]].copy()
        sub = sub.rename(columns={col: "data"})
        sub["variable"] = col

        # 按 group 分组计算 CDF（rank / n）
        sub["CDF"] = sub.groupby(group_col)["data"].transform(empirical_cdf)

        # Weibull 变换：ln(-ln(1 - median_rank))
        sub["mr"] = sub.groupby(group_col)["data"].transform(median_rank)
        sub["weibull"] = weibull_transform(sub["mr"])
        sub = sub.drop(columns=["mr"])

        # limit：按测试项名称查 limit_map
        lim = get_limit(col, limit_map)
        sub["limit"] = lim if lim is not None else None

        frames.append(sub)

    result = pd.concat(frames, ignore_index=True)
    result = result[[id_col, group_col, "variable", "data", "CDF", "weibull", "limit"]]
    # 统一列名为 "id" 和 "group"，方便下游模块使用
    result = result.rename(columns={id_col: "id", group_col: "group"})
    return result


# ──────────────────────────────────────────────
# 统计汇总
# ──────────────────────────────────────────────


def build_summary(
    df: pd.DataFrame,
    group_col: str,
    data_cols: list[str],
    limit_map: dict[str, float] | None = None,
) -> pd.DataFrame:
    """
    构建统计汇总 DataFrame。

    返回列: 测试项, 总模块数, Group, 均值, 标准差, 25%分位, 75%分位, 中位数,
            最小值, 最大值, 变异系数(CV%), Weibull β, Weibull η, 拟合 R², Limit, limit处 CDF
    """
    import numpy as np

    rows = []

    for col in data_cols:
        # 查该测试项的 limit 值
        limit_val = get_limit(col, limit_map)

        for grp_name, grp_df in df.groupby(group_col):
            values = grp_df[col].dropna()
            if len(values) == 0:
                continue

            # Weibull 拟合
            mr = median_rank(values)
            wb_fit = fit_weibull(values, mr)
            beta, eta, r2 = wb_fit if wb_fit else (None, None, None)

            # limit 处的 CDF: F(limit) = 1 - exp(-(limit/eta)^beta)
            cdf_at_limit = None
            if beta is not None and eta is not None and limit_val is not None:
                try:
                    cdf_at_limit = 1 - np.exp(-((limit_val / eta) ** beta))
                except (OverflowError, ZeroDivisionError):
                    cdf_at_limit = None

            rows.append(
                {
                    "测试项": col,
                    "总模块数": len(values),
                    "Group": grp_name,
                    "均值": round(values.mean(), 4),
                    "标准差": round(values.std(ddof=0), 4),
                    "25%分位": round(values.quantile(0.25), 4),
                    "75%分位": round(values.quantile(0.75), 4),
                    "中位数": round(values.median(), 4),
                    "最小值": round(values.min(), 4),
                    "最大值": round(values.max(), 4),
                    "变异系数(CV%)": round(
                        (values.std(ddof=0) / values.mean() * 100)
                        if values.mean() != 0
                        else float("nan"),
                        2,
                    ),
                    "Weibull β": beta,
                    "Weibull η": eta,
                    "拟合 R²": r2,
                    "Limit": limit_val,
                    "limit处 CDF": cdf_at_limit,
                }
            )

    if not rows:
        return pd.DataFrame(columns=SUMMARY_HEADERS)

    result = pd.DataFrame(rows)
    return result[SUMMARY_HEADERS]


# ──────────────────────────────────────────────
# Excel 写出
# ──────────────────────────────────────────────


def _apply_header_style(ws, row_idx: int, col_count: int) -> None:
    """给指定行的表头单元格应用样式。"""
    for c in range(1, col_count + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER


def _apply_data_border(ws, start_row: int, end_row: int, col_count: int) -> None:
    """给数据区域加边框。"""
    for r in range(start_row, end_row + 1):
        for c in range(1, col_count + 1):
            ws.cell(row=r, column=c).border = THIN_BORDER
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")


def write_summary_sheet(
    wb: Workbook,
    summary_df: pd.DataFrame,
    raw_df: pd.DataFrame | None = None,
    params: dict | None = None,
) -> None:
    """将参数信息 + 统计汇总 + 原始数据 写入第一个 sheet。"""

    PARAM_LABEL_FONT = Font(name="微软雅黑", bold=True, size=10)
    PARAM_VALUE_FONT = Font(name="微软雅黑", size=10)

    ws = wb.active
    ws.title = "统计汇总"

    summary_start_row = 1

    # ── 参数信息（可选，放在最顶部） ──
    if params:
        for i, (label, value) in enumerate(params.items()):
            r = 1 + i
            ws.cell(row=r, column=1, value=label).font = PARAM_LABEL_FONT
            ws.cell(row=r, column=2, value=str(value)).font = PARAM_VALUE_FONT
        # 参数区列宽
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 48

        summary_start_row = len(params) + 2  # 空一行

    # ── 统计汇总分区标题（合并居中） ──
    SECTION_TITLE_FONT = Font(name="微软雅黑", bold=True, size=12, color="4472C4")
    ws.merge_cells(
        start_row=summary_start_row, start_column=1,
        end_row=summary_start_row, end_column=len(SUMMARY_HEADERS),
    )
    cell = ws.cell(row=summary_start_row, column=1, value="统计汇总")
    cell.font = SECTION_TITLE_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    summary_start_row += 1  # 标题占一行

    # ── 统计汇总表 ──
    hdr_row = summary_start_row
    for col_idx, header in enumerate(SUMMARY_HEADERS, start=1):
        ws.cell(row=hdr_row, column=col_idx, value=header)

    for row_idx, (_, row) in enumerate(summary_df.iterrows(), start=hdr_row + 1):
        for col_idx, header in enumerate(SUMMARY_HEADERS, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row[header])

    _apply_header_style(ws, hdr_row, len(SUMMARY_HEADERS))
    if len(summary_df) > 0:
        # limit处 CDF(%) 列用科学计数法
        cdf_col = len(SUMMARY_HEADERS)  # 最后一列
        for row_idx in range(hdr_row + 1, hdr_row + 1 + len(summary_df)):
            ws.cell(row=row_idx, column=cdf_col).number_format = '0.00E+00'
        _apply_data_border(
            ws, hdr_row + 1, hdr_row + len(summary_df), len(SUMMARY_HEADERS)
        )

    # ── 原始数据（接在统计汇总后面，中间空一行） ──
    raw_title_row = hdr_row + len(summary_df) + 2  # +1 统计表头, +1 spacer
    raw_start_row = raw_title_row + 1  # 标题下面才是表头

    if raw_df is not None and not raw_df.empty:
        raw_headers = list(raw_df.columns)
        n_raw_cols = len(raw_headers)

        # 原始数据分区标题（合并居中）
        ws.merge_cells(
            start_row=raw_title_row, start_column=1,
            end_row=raw_title_row, end_column=n_raw_cols,
        )
        cell = ws.cell(row=raw_title_row, column=1, value="原始数据")
        cell.font = SECTION_TITLE_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_idx, header in enumerate(raw_headers, start=1):
            ws.cell(row=raw_start_row, column=col_idx, value=header)

        _apply_header_style(ws, raw_start_row, n_raw_cols)

        for ri, (_, row) in enumerate(raw_df.iterrows(), start=raw_start_row + 1):
            for ci, header in enumerate(raw_headers, start=1):
                ws.cell(row=ri, column=ci, value=row[header])

        _apply_data_border(
            ws, raw_start_row + 1, raw_start_row + len(raw_df), n_raw_cols
        )

    # 自动列宽
    for col_idx in range(1, len(SUMMARY_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14


def write_data_sheets(
    wb: Workbook,
    long_df: pd.DataFrame,
    data_cols: list[str],
) -> None:
    """
    将长表数据按 variable 分组写出到各自的 sheet。

    每个 sheet 列: id, group, data, CDF, weibull, limit
    """
    sheet_columns = ["id", "group", "data", "CDF", "weibull", "limit"]
    sheet_headers = ["id", "group", "数据", "CDF", "weibull", "limit"]
    n_cols = len(sheet_headers)

    for col_name in data_cols:
        sub = long_df[long_df["variable"] == col_name][sheet_columns].copy()
        # 按 CDF 排序（从小到大，稍后用 data 排更直观）
        sub = sub.sort_values("data")

        ws = wb.create_sheet(title=col_name[:31])  # sheet 名最长 31 字符

        # 写表头
        for ci, h in enumerate(sheet_headers, start=1):
            ws.cell(row=1, column=ci, value=h)

        # 写数据
        for ri, (_, row) in enumerate(sub.iterrows(), start=2):
            ws.cell(row=ri, column=1, value=row["id"])
            ws.cell(row=ri, column=2, value=row["group"])
            ws.cell(row=ri, column=3, value=row["data"])
            ws.cell(row=ri, column=4, value=row["CDF"])
            ws.cell(row=ri, column=5, value=row["weibull"])
            ws.cell(row=ri, column=6, value=row["limit"])

        # 样式
        _apply_header_style(ws, 1, n_cols)
        if len(sub) > 0:
            _apply_data_border(ws, 2, 1 + len(sub), n_cols)

        # 自动列宽
        widths = [14, 10, 12, 12, 14, 10]
        for ci, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(ci)].width = w


# ──────────────────────────────────────────────
# Excel 内嵌图表
# ──────────────────────────────────────────────


def add_excel_chart(
    wb: Workbook,
    df: pd.DataFrame,
    id_col: str,
    group_col: str,
    data_cols: list[str],
    limit_map: dict[str, float] | None = None,
    x_axis: str = "data",
    y_axis: str = "CDF",
    x_scale: str = "linear",
    y_scale: str = "linear",
    show_limit: bool = True,
    chart_width: float = 20,
    chart_height: float = 12,
    auto_axis: bool = True,
    x_min: float | None = None,
    x_max: float | None = None,
    y_min: float | None = None,
    y_max: float | None = None,
    marker_size: int = 5,
    chart_title: str | None = None,
    x_label: str | None = None,
    y_label: str | None = None,
) -> Workbook:
    """
    在 Workbook 的每个数据 sheet 中绘制散点图。

    Parameters
    ----------
    wb : Workbook
        已含数据 sheet 的 openpyxl Workbook（由 write_data_sheets 写出）。
    df : DataFrame
        原始宽表数据，含 id_col, group_col, data_cols。
    id_col : str
        ID 列名。
    group_col : str
        分组列名。
    data_cols : list[str]
        数据列名列表。
    limit_map : dict | None
        {r"测试项正则": limit_value} 用于绘制 limit 参考线。
    x_axis : str
        X 轴数据列: "data"
    y_axis : str
        Y 轴数据列: "CDF" | "weibull"
    x_scale : str
        X 轴刻度: "linear" | "log"
    y_scale : str
        Y 轴刻度: "linear" | "log"
    show_limit : bool
        是否绘制 limit 垂直线。
    chart_width : float
        图表宽度 (cm)。
    chart_height : float
        图表高度 (cm)。
    auto_axis : bool
        True: 自动设置轴范围; False: 不手动干预。
    marker_size : int
        散点 marker 大小。
    chart_title : str | None
        图表标题。
    x_label : str | None
        X 轴标签。
    y_label : str | None
        Y 轴标签。

    Returns
    -------
    Workbook
        传入的 wb（已添加图表），方便链式调用。
    """
    import math

    import numpy as np

    # --- 先计算长表，获取 CDF/weibull/limit ---
    long_df = compute_statistics(df, id_col, group_col, data_cols, limit_map)

    # --- 再更新各个 data sheet 的数据（追加图表） ---
    # 但 write_data_sheets 已经写过了，这里需要重新把 CDF/weibull/limit 写到对应位置
    # 所以采用: 先清空数据 sheet，重写，再加图

    for col_name in data_cols:
        ws = wb[col_name[:31]]

        sub = long_df[long_df["variable"] == col_name].sort_values("data")
        if sub.empty:
            continue

        groups = sub["group"].unique().tolist()
        n_rows = len(sub)
        # 数据结束行 = 表头(1) + n_rows
        data_end = 1 + n_rows

        # --- 构建散点图 ---
        chart = ScatterChart()
        chart.width = chart_width
        chart.height = chart_height

        # 坐标轴缩放类型
        if x_scale == "log":
            chart.x_axis.scaling.logBase = 10
        if y_scale == "log":
            chart.y_axis.scaling.logBase = 10

        # 为每个 group 建一个系列
        # 需要找到该 group 数据在 sheet 中的行号范围
        for gi, grp_name in enumerate(groups):
            grp_mask = sub["group"] == grp_name
            grp_data = sub[grp_mask]
            grp_rows = grp_data.index.tolist()
            if not grp_rows:
                continue

            # group 数据在 sheet 中的行号（2-based: 第1行是表头）
            first_row = grp_data.index[0] + 2  # sub 的 index 是原 df index
            # 实际上 sub.sort_values 后索引变了，需要重新定位
            # 更好的方式：按组内行号定位

            # --- 重新定位：sub 排序后索引重置 ---
            # 因为 sub 是排序后的，行号已重置
            # 我们改用另一种方式：不算小组偏移，直接用 sub 的行位置

        # 上面逻辑有问题——sub 排序后索引乱序。改用：
        # 先对 sub 重置索引，然后按组写回 sheet、同时记录行范围

        # 重新写数据 sheet，按组排序
        sub = sub.reset_index(drop=True)
        sub = sub.sort_values(["group", "data"])

        # 重新写数据
        sheet_columns = ["id", "group", "data", "CDF", "weibull", "limit"]
        for ri in range(len(sub)):
            row_data = sub.iloc[ri]
            ws.cell(row=2 + ri, column=1, value=row_data["id"])
            ws.cell(row=2 + ri, column=2, value=row_data["group"])
            ws.cell(row=2 + ri, column=3, value=row_data["data"])
            ws.cell(row=2 + ri, column=4, value=row_data["CDF"])
            ws.cell(row=2 + ri, column=5, value=row_data["weibull"])
            ws.cell(row=2 + ri, column=6, value=row_data["limit"])

        # --- 确定 x_col, y_col 在 sheet 中的列号 ---
        x_col_map = {
            "id": 1,
            "group": 2,
            "data": 3,
            "CDF": 4,
            "weibull": 5,
            "limit": 6,
        }
        x_col_num = x_col_map.get(x_axis, 3)
        y_col_num = x_col_map.get(y_axis, 4)

        # --- 每个 group 一个系列 ---
        for gi, grp_name in enumerate(groups):
            grp_mask = sub["group"] == grp_name
            grp_indices = sub[grp_mask].index.tolist()
            if not grp_indices:
                continue

            first = 2 + grp_indices[0]
            last = 2 + grp_indices[-1]

            x_values = Reference(ws, min_col=x_col_num, min_row=first, max_row=last)
            y_values = Reference(ws, min_col=y_col_num, min_row=first, max_row=last)

            series = Series(y_values, x_values, title=str(grp_name))

            # marker 样式
            color = COLOR_PALETTE[gi % len(COLOR_PALETTE)]
            symbol = MARKER_SYMBOLS[gi % len(MARKER_SYMBOLS)]
            series.marker.symbol = symbol
            series.marker.size = marker_size
            series.marker.graphicalProperties.solidFill = color
            series.graphicalProperties.line.noFill = True  # 无连线

            chart.series.append(series)

        # --- limit 参考线（可选） ---
        if show_limit and limit_map:
            limit_values = set()
            for _, row in sub.iterrows():
                lim = row["limit"]
                if not (lim is None or (isinstance(lim, float) and np.isnan(lim))):
                    limit_values.add(lim)

            for lim in sorted(limit_values):
                # 垂直线：两个点 (lim, y_min) 和 (lim, y_max)
                # y_min, y_max 从数据中取
                y_all = sub[y_axis].dropna()
                if y_all.empty:
                    y_all = pd.Series([0, 1])

                y_min_val = float(y_all.min())
                y_max_val = float(y_all.max())

                # 在 sheet 中写入两个点，放在数据后面
                lim_start_row = data_end + 1
                lim_end_row = data_end + 2

                ws.cell(row=lim_start_row, column=x_col_num, value=lim)
                ws.cell(row=lim_end_row, column=x_col_num, value=lim)
                ws.cell(row=lim_start_row, column=y_col_num, value=y_min_val)
                ws.cell(row=lim_end_row, column=y_col_num, value=y_max_val)

                x_lim = Reference(
                    ws,
                    min_col=x_col_num,
                    min_row=lim_start_row,
                    max_row=lim_end_row,
                )
                y_lim = Reference(
                    ws,
                    min_col=y_col_num,
                    min_row=lim_start_row,
                    max_row=lim_end_row,
                )

                lim_series = Series(y_lim, x_lim, title=f"limit={lim}")
                lim_series.marker.symbol = "none"  # 无 marker
                lim_series.graphicalProperties.line.solidFill = "FF0000"  # 红色
                lim_series.graphicalProperties.line.width = 20000  # EMU, ~1.5pt
                lim_series.graphicalProperties.line.dashStyle = "dash"  # 虚线

                chart.series.append(lim_series)

                data_end += 2  # 更新数据结束位置

        # --- 轴范围 ---
        if not auto_axis:
            # 优先使用用户指定的强制范围，否则从数据取
            _x_min = x_min if x_min is not None else (
                float(sub[x_axis].dropna().min()) if not sub[x_axis].dropna().empty else None
            )
            _x_max = x_max if x_max is not None else (
                float(sub[x_axis].dropna().max()) if not sub[x_axis].dropna().empty else None
            )
            _y_min = y_min if y_min is not None else (
                float(sub[y_axis].dropna().min()) if not sub[y_axis].dropna().empty else None
            )
            _y_max = y_max if y_max is not None else (
                float(sub[y_axis].dropna().max()) if not sub[y_axis].dropna().empty else None
            )

            if _x_min is not None and _x_max is not None:
                padding = (_x_max - _x_min) * 0.05 if _x_max > _x_min else 1.0
                chart.x_axis.scaling.min = round(_x_min - padding, 4)
                chart.x_axis.scaling.max = round(_x_max + padding, 4)

            if _y_min is not None and _y_max is not None:
                padding = (_y_max - _y_min) * 0.05 if _y_max > _y_min else 0.1
                chart.y_axis.scaling.min = round(_y_min - padding, 4)
                chart.y_axis.scaling.max = round(_y_max + padding, 4)

        # 网格线 + 次级刻度 + 轴显示
        _add_gridlines(chart)
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.minorTickMark = "out"
        chart.y_axis.minorTickMark = "out"
        chart.x_axis.tickLblSkip = 1
        chart.y_axis.tickLblSkip = 1
        chart.x_axis.tickLblPos = "low"
        chart.y_axis.tickLblPos = "low"
        # 数值格式
        chart.x_axis.numFmt = '0.0###'
        chart.y_axis.numFmt = '0.0###'

        # 标题 + 图例右侧竖排
        chart.title = chart_title or f"{col_name} CDF 分布"
        chart.title.overlay = False
        chart.legend.position = "r"
        chart.legend.overlay = False
        chart.x_axis.title = x_label or col_name
        chart.y_axis.title = y_label or ("CDF" if y_axis == "CDF" else "ln(-ln(1-MR))")
        chart.x_axis.title.overlay = False
        chart.y_axis.title.overlay = False

        chart.plot_area.layout = Layout(
        manualLayout=ManualLayout(
            xMode="factor", x=0.12,
            yMode="factor", y=0.15,
            wMode="factor", w=0.72,
            hMode="factor", h=0.70,
        )
    )

        # 添加图表到 sheet
        # 放在数据表右侧
        chart_col_letter = get_column_letter(8)  # H 列
        ws.add_chart(chart, f"{chart_col_letter}2")

    return wb


# ──────────────────────────────────────────────
# 数据读取
# ──────────────────────────────────────────────


def read_file(filepath: str | Path) -> pd.DataFrame:
    """
    读取 Excel 文件返回 DataFrame。

    初版仅支持 .xlsx。后续可扩展 CSV 等。
    """
    filepath = Path(filepath)
    if filepath.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError(f"不支持的文件类型: {filepath.suffix}，仅支持 .xlsx / .xlsm")
    return pd.read_excel(filepath)


# ──────────────────────────────────────────────
# 一键流水线
# ──────────────────────────────────────────────


def process(
    df: pd.DataFrame,
    id_col: str,
    group_col: str,
    data_cols: list[str],
    output_path: str | Path,
    limit_map: dict[str, float] | None = None,
    x_axis: str = "data",
    y_axis: str = "CDF",
    x_scale: str = "linear",
    y_scale: str = "linear",
    show_limit: bool = True,
    chart_width: float = 20,
    chart_height: float = 12,
    auto_axis: bool = True,
    x_min: float | None = None,
    x_max: float | None = None,
    y_min: float | None = None,
    y_max: float | None = None,
    marker_size: int = 5,
    chart_title: str | None = None,
    x_label: str | None = None,
    y_label: str | None = None,
) -> Path:
    """
    一站式处理流水线：算 → 写 → 图 → 保存。

    Parameters
    ----------
    df : DataFrame
        输入数据。需含 id_col, group_col, data_cols 列。
    id_col : str
        样品 ID 列名。
    group_col : str
        分组列名。
    data_cols : list[str]
        要处理的数据列名。
    output_path : str | Path
        输出 Excel 路径。
    limit_map : dict | None
        {regex_pattern: limit_value} 用于 limit 参考线。
    x_axis, y_axis : str
        X/Y 轴数据列 ("data", "CDF", "weibull" 等)。
    x_scale, y_scale : str
        坐标轴缩放 ("linear" | "log")。
    show_limit : bool
        是否绘制 limit 线。
    chart_width, chart_height : float
        图表尺寸 (cm)。
    auto_axis : bool
        True 自动轴范围。
    marker_size : int
        散点 marker 大小。
    chart_title, x_label, y_label : str | None
        图表标题及轴标签。

    Returns
    -------
    Path
        输出文件路径。
    """
    # 1. 计算
    long_df = compute_statistics(df, id_col, group_col, data_cols, limit_map)
    summary_df = build_summary(df, group_col, data_cols, limit_map=limit_map)

    # 3. 写 Workbook
    wb = Workbook()
    write_summary_sheet(
        wb,
        summary_df,
        raw_df=df,
        params={
            "数据来源": f"DataFrame ({len(df)} 行 × {len(df.columns)} 列)",
            "输出文件": str(output_path),
            "ID 列": id_col,
            "分组列": group_col,
            "数据列": ", ".join(data_cols),
            "X 轴": f"{x_axis}（默认）" if x_axis == "data" else x_axis,
            "Y 轴": f"{y_axis}（默认）" if y_axis == "CDF" else y_axis,
            "X 缩放": f"{x_scale}（默认）" if x_scale == "linear" else x_scale,
            "Y 缩放": f"{y_scale}（默认）" if y_scale == "linear" else y_scale,
            "显示 Limit": "是（默认）" if show_limit else "否",
            "Limit Map": str(limit_map) if limit_map else "无（默认）",
            "自动轴范围": f"{'是（默认）' if auto_axis else '否'}",
            "X 范围": f"{x_min}~{x_max}" if x_min is not None or x_max is not None else "自动（默认）",
            "Y 范围": f"{y_min}~{y_max}" if y_min is not None or y_max is not None else "自动（默认）",
            "图表宽度": f"{chart_width} cm（默认）" if chart_width == 20 else f"{chart_width} cm",
            "图表高度": f"{chart_height} cm（默认）" if chart_height == 12 else f"{chart_height} cm",
            "Marker 大小": f"{marker_size}（默认）" if marker_size == 5 else str(marker_size),
            "图表标题": f"{chart_title or '自动（默认）'}",
            "X 轴标签": f"{x_label or '自动（默认）'}",
            "Y 轴标签": f"{y_label or '自动（默认）'}",
        },
    )
    write_data_sheets(wb, long_df, data_cols)

    # 4. 给每个数据 sheet 添加散点图
    add_excel_chart(
        wb=wb,
        df=df,
        id_col=id_col,
        group_col=group_col,
        data_cols=data_cols,
        limit_map=limit_map,
        x_axis=x_axis,
        y_axis=y_axis,
        x_scale=x_scale,
        y_scale=y_scale,
        show_limit=show_limit,
        chart_width=chart_width,
        chart_height=chart_height,
        auto_axis=auto_axis,
        x_min=x_min,
        x_max=x_max,
        y_min=y_min,
        y_max=y_max,
        marker_size=marker_size,
        chart_title=chart_title,
        x_label=x_label,
        y_label=y_label,
    )

    # 5. 保存
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    return output_path


if __name__ == "__main__":
    # 测试运行
    test_input = Path("test_data.xlsx")
    test_output = Path("test_output.xlsx")
    process(
        input_path=test_input,
        id_col="id",
        group_col="group",
        data_cols=["Vth", "BV"],
        output_path=test_output,
        limit_map={r"G1": 3.5},
        x_axis="data",
        y_axis="CDF",
        x_scale="linear",
        y_scale="linear",
        show_limit=True,
        chart_width=20,
        chart_height=12,
        auto_axis=False,
        marker_size=5,
        chart_title="Reliability Statistics",
        x_label="Data Value",
        y_label="CDF",
    )