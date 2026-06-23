"""reli_stat 单元测试"""

import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from reli_stat import (
    compute_statistics,
    get_limit,
    empirical_cdf,
    median_rank,
    weibull_transform,
    build_summary,
    process,
)


# ── empirical_cdf ──────────────────────────────


def test_empirical_cdf_single_group():
    values = pd.Series([2.0, 2.5, 2.3, 2.8, 2.1])
    cdf = empirical_cdf(values)

    assert len(cdf) == 5
    assert (cdf > 0).all()
    assert (cdf <= 1).all()
    sorted_cdf = cdf.loc[values.sort_values().index]
    assert sorted_cdf.is_monotonic_increasing


def test_empirical_cdf_ties():
    values = pd.Series([1.0, 2.0, 2.0, 3.0])
    cdf = empirical_cdf(values)
    assert cdf.iloc[1] == cdf.iloc[2]


def test_empirical_cdf_single_value():
    values = pd.Series([5.0])
    cdf = empirical_cdf(values)
    assert cdf.iloc[0] == 1.0


def test_empirical_cdf_max_n():
    values = pd.Series([1.0, 2.0, 3.0])
    cdf = empirical_cdf(values)
    assert cdf.max() == 1.0


# ── median_rank ────────────────────────────────


def test_median_rank_basic():
    values = pd.Series([2.1, 2.5, 2.3, 2.8, 2.0])
    mr = median_rank(values)
    assert len(mr) == 5
    assert (mr > 0).all()
    assert (mr < 1).all()


def test_median_rank_single():
    mr = median_rank(pd.Series([5.0]))
    assert abs(mr.iloc[0] - 0.5) < 1e-9


# ── weibull_transform ──────────────────────────


def test_weibull_basic():
    mr = pd.Series([0.1, 0.3, 0.5, 0.7, 0.9])
    w = weibull_transform(mr)
    assert w.is_monotonic_increasing
    assert not w.isna().any()


def test_weibull_boundary():
    mr = pd.Series([0.0, 0.5, 1.0])
    w = weibull_transform(mr)
    assert pd.isna(w.iloc[0])
    assert pd.isna(w.iloc[2])
    assert not pd.isna(w.iloc[1])


# ── get_limit ──────────────────────────────────


def test_get_limit_hit():
    """精确匹配测试项名。"""
    assert get_limit("Vth", {"Vth": 3.0, "BVdss": 650}) == 3.0


def test_get_limit_regex():
    """正则匹配：r\"Vth.*\" 可匹配 Vth_25C, Vth_150C 等变体。"""
    assert get_limit("Vth_25C", {r"Vth.*": 3.0}) == 3.0
    assert get_limit("Vth_150C", {r"Vth.*": 3.0}) == 3.0


def test_get_limit_first_match():
    """多个 pattern 命中时取第一个。"""
    assert get_limit("Vth_hot", {r"Vth": 3.0, r"Vth.*": 2.5}) == 3.0


def test_get_limit_miss():
    """未命中返回 None。"""
    assert get_limit("Rds_on", {"Vth": 3.0}) is None


def test_get_limit_empty_map():
    """空 map 返回 None。"""
    assert get_limit("Vth", None) is None
    assert get_limit("Vth", {}) is None


# ── compute_statistics ─────────────────────────


def test_compute_statistics():
    df = pd.DataFrame(
        {
            "id": ["G1-1", "G1-2", "G2-1", "G2-2"],
            "group": ["A", "A", "B", "B"],
            "Vth": [2.0, 3.0, 2.5, 2.8],
        }
    )

    result = compute_statistics(
        df,
        id_col="id",
        group_col="group",
        data_cols=["Vth"],
        limit_map={"Vth": 3.5},
    )

    assert list(result.columns) == [
        "id", "group", "variable", "data", "CDF", "weibull", "limit",
    ]
    assert len(result) == 4

    # Vth 的 limit 应该是 3.5（所有行一样）
    assert (result["limit"] == 3.5).all()

    # G1 的 CDF: rank=1,n=2→0.5, rank=2,n=2→1.0
    g1 = result[result["group"] == "A"].sort_values("data")
    assert g1["CDF"].iloc[0] == 0.5
    assert g1["CDF"].iloc[1] == 1.0


def test_compute_statistics_limit_per_var():
    """不同测试项可以有不同 limit。"""
    df = pd.DataFrame(
        {
            "id": ["X1", "X2"],
            "group": ["A", "A"],
            "Vth": [1.0, 2.0],
            "BV": [100, 200],
        }
    )

    result = compute_statistics(
        df, "id", "group", ["Vth", "BV"],
        limit_map={"Vth": 3.0, "BV": 500},
    )

    vth_rows = result[result["variable"] == "Vth"]
    bv_rows = result[result["variable"] == "BV"]
    assert (vth_rows["limit"] == 3.0).all()
    assert (bv_rows["limit"] == 500).all()


def test_compute_statistics_multi_col():
    df = pd.DataFrame(
        {
            "id": ["X1", "X2"],
            "group": ["A", "A"],
            "Vth": [1.0, 2.0],
            "BV": [100, 200],
        }
    )

    result = compute_statistics(df, "id", "group", ["Vth", "BV"])
    assert len(result) == 4
    assert set(result["variable"]) == {"Vth", "BV"}


# ── build_summary ──────────────────────────────


def test_build_summary():
    df = pd.DataFrame(
        {
            "id": ["A1", "A2", "A3", "B1", "B2"],
            "group": ["G1", "G1", "G1", "G2", "G2"],
            "Vth": [2.0, 3.0, 4.0, 2.5, 2.7],
        }
    )

    summary = build_summary(df, "group", ["Vth"])

    assert len(summary) == 2
    assert list(summary["测试项"]) == ["Vth", "Vth"]
    assert summary.iloc[0]["总模块数"] == 3
    assert round(summary.iloc[0]["均值"], 1) == 3.0
    # Weibull 拟合列存在
    assert "Weibull β" in summary.columns
    assert "Weibull η" in summary.columns
    assert "拟合 R²" in summary.columns
    # 3 个点的 Weibull 拟合应有效
    assert summary.iloc[0]["拟合 R²"] is not None


# ── 集成测试：端到端 pipeline ──────────────────

FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE_XLSX = FIXTURES / "sample.xlsx"
OUTPUT_XLSX = FIXTURES / "output.xlsx"
OUTPUT_NOLIMIT_XLSX = FIXTURES / "output_nolimit.xlsx"


def test_pipeline_end_to_end():
    """完整流水线：读 sample.xlsx → 计算 → 写 Excel → 验证输出。

    每次运行刷新 output.xlsx，包含所有参数和 limit 线。"""
    import openpyxl

    output = process(
        input_path=SAMPLE_XLSX,
        id_col="样品编号",
        group_col="批次",
        data_cols=["Vth", "BVdss", "Rds_on"],
        output_path=OUTPUT_XLSX,
        limit_map={r"Vth": 3.0, r"BVdss": 660, r"Rds_on": 2.0},
        x_axis="data",
        y_axis="CDF",
        show_limit=True,
        chart_width=18,
        chart_height=10,
        marker_size=6,
        x_scale="log",
        y_scale="log",
    )
    assert output == OUTPUT_XLSX
    assert OUTPUT_XLSX.exists()

    wb = openpyxl.load_workbook(OUTPUT_XLSX)

    # --- 验证 sheet 结构 ---
    assert wb.sheetnames == ["统计汇总", "Vth", "BVdss", "Rds_on"]

    # 统计汇总
    ws = wb["统计汇总"]
    # 参数区应在顶部（20 个参数）
    assert ws.cell(1, 1).value == "输入文件"
    # 20 params + 空行 + 标题"统计汇总" → 统计表头行23
    stat_header_row = 23
    assert ws.cell(stat_header_row, 1).value == "测试项"
    assert ws.cell(stat_header_row + 1, 1).value == "Vth"
    # 16 列（含 Limit / limit处 CDF）
    assert ws.cell(stat_header_row, 15).value == "Limit"
    assert ws.cell(stat_header_row, 16).value == "limit处 CDF(%)"
    # 分区标题
    assert ws.cell(22, 1).value == "统计汇总"  # 合并单元格后该值在第一个格子
    # 原始数据：22(标题) + 1(统计头) + 9(数据) + 1(空) + 1(原始标题) + 1 = 35
    raw_hdr_row = 35
    assert ws.cell(raw_hdr_row, 1).value == "样品编号"

    # --- 验证数据 sheet 有图表和数据 ---
    for sheet_name in ["Vth", "BVdss", "Rds_on"]:
        ws = wb[sheet_name]
        assert ws.max_row > 2  # 有数据行
        assert len(ws._charts) == 1

        chart = ws._charts[0]
        # 检查标题、轴标签已设置
        assert chart.title is not None
        assert chart.x_axis.title is not None
        assert chart.y_axis.title is not None
        # 3 组 + limit 线
        assert len(chart.series) >= 4, f"{sheet_name}: expected ≥4 series, got {len(chart.series)}"

    # --- 验证 CDF 值在 (0, 1] 且每组内单调 ---
    ws_vth = wb["Vth"]
    rows = []
    for r in range(2, ws_vth.max_row + 1):
        g = ws_vth.cell(r, 2).value
        d = ws_vth.cell(r, 3).value
        c = ws_vth.cell(r, 4).value
        if all(isinstance(x, (int, float)) for x in (d, c)):
            rows.append((g, d, c))

    for grp in set(g for g, _, _ in rows):
        grp_cdf = [c for g, _, c in sorted(
            [(g, d, c) for g, d, c in rows if g == grp], key=lambda x: x[1]
        )]
        assert all(grp_cdf[i] <= grp_cdf[i + 1] for i in range(len(grp_cdf) - 1)), \
            f"CDF not monotonic within group {grp}"


def test_pipeline_no_limit():
    """不传 limit_map 时仍正常运行（用独立输出文件）。"""
    output = process(
        input_path=SAMPLE_XLSX,
        id_col="样品编号",
        group_col="批次",
        data_cols=["Vth"],
        output_path=OUTPUT_NOLIMIT_XLSX,
        show_limit=False,
    )
    assert OUTPUT_NOLIMIT_XLSX.exists()

    import openpyxl
    wb = openpyxl.load_workbook(OUTPUT_NOLIMIT_XLSX)
    assert wb.sheetnames == ["统计汇总", "Vth"]
    chart = wb["Vth"]._charts[0]
    # 3 组，无 limit
    assert len(chart.series) == 3


def test_pipeline_defaults():
    """默认参数输出（备注：默认）。"""
    out = FIXTURES / "output_default.xlsx"
    process(
        input_path=SAMPLE_XLSX,
        id_col="样品编号",
        group_col="批次",
        data_cols=["Vth", "BVdss", "Rds_on"],
        output_path=out,
        chart_title="默认参数",
        x_label="数据",
        y_label="CDF",
    )
    assert out.exists()


# ── 多尺度组合输出 ──

SCALE_COMBOS = [
    ("linear", "linear"),
    ("log", "linear"),
    ("linear", "log"),
    ("log", "log"),
]


@pytest.mark.parametrize("x_scale,y_scale", SCALE_COMBOS)
def test_scale_combos(x_scale: str, y_scale: str):
    """不同 x/y 缩放组合输出独立文件到 Windows。"""
    out = FIXTURES / f"output_x{x_scale}_y{y_scale}.xlsx"
    result = process(
        input_path=SAMPLE_XLSX,
        id_col="样品编号",
        group_col="批次",
        data_cols=["Vth", "BVdss", "Rds_on"],
        output_path=out,
        limit_map={r"Vth": 3.0, r"BVdss": 660, r"Rds_on": 2.0},
        x_scale=x_scale,
        y_scale=y_scale,
        show_limit=True,
        chart_width=18,
        chart_height=10,
    )
    assert out.exists()
    print(f"  ✓ {out}")

    # 快速校验
    import openpyxl
    wb = openpyxl.load_workbook(out)
    assert len(wb.sheetnames) == 4
    for sn in ["Vth", "BVdss", "Rds_on"]:
        ch = wb[sn]._charts[0]
        assert ch.x_axis.scaling.logBase == (10 if x_scale == "log" else None)
        assert ch.y_axis.scaling.logBase == (10 if y_scale == "log" else None)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])