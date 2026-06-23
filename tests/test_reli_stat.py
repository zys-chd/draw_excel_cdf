"""reli_stat 单元测试 + 集成测试"""

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
    read_file,
)


FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE_XLSX = FIXTURES / "sample.xlsx"
OUTPUT_XLSX = FIXTURES / "output.xlsx"
OUTPUT_NOLIMIT_XLSX = FIXTURES / "output_nolimit.xlsx"


def _load_sample():
    """加载测试数据。"""
    return read_file(SAMPLE_XLSX)


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
    assert get_limit("Vth", {"Vth": 3.0, "BVdss": 650}) == 3.0


def test_get_limit_regex():
    assert get_limit("Vth_25C", {r"Vth.*": 3.0}) == 3.0
    assert get_limit("Vth_150C", {r"Vth.*": 3.0}) == 3.0


def test_get_limit_first_match():
    assert get_limit("Vth_hot", {r"Vth": 3.0, r"Vth.*": 2.5}) == 3.0


def test_get_limit_miss():
    assert get_limit("Rds_on", {"Vth": 3.0}) is None


def test_get_limit_empty_map():
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

    result = compute_statistics(df, id_col="id", group_col="group",
                                data_cols=["Vth"], limit_map={"Vth": 3.5})
    assert len(result) == 4
    assert (result["limit"] == 3.5).all()


def test_compute_statistics_limit_per_var():
    df = pd.DataFrame(
        {"id": ["X1", "X2"], "group": ["A", "A"], "Vth": [1.0, 2.0], "BV": [100, 200]}
    )
    result = compute_statistics(df, "id", "group", ["Vth", "BV"],
                                limit_map={"Vth": 3.0, "BV": 500})
    assert (result[result["variable"] == "Vth"]["limit"] == 3.0).all()
    assert (result[result["variable"] == "BV"]["limit"] == 500).all()


def test_compute_statistics_multi_col():
    df = pd.DataFrame(
        {"id": ["X1", "X2"], "group": ["A", "A"], "Vth": [1.0, 2.0], "BV": [100, 200]}
    )
    result = compute_statistics(df, "id", "group", ["Vth", "BV"])
    assert len(result) == 4


# ── build_summary ──────────────────────────────


def test_build_summary():
    df = pd.DataFrame(
        {"id": ["A1", "A2", "A3", "B1", "B2"], "group": ["G1", "G1", "G1", "G2", "G2"],
         "Vth": [2.0, 3.0, 4.0, 2.5, 2.7]}
    )
    summary = build_summary(df, "group", ["Vth"])
    assert len(summary) == 2
    assert "Weibull β" in summary.columns
    assert summary.iloc[0]["拟合 R²"] is not None


# ── 集成测试 ───────────────────────────────────


def test_pipeline_end_to_end():
    import openpyxl
    df = _load_sample()

    output = process(
        df=df,
        id_col="样品编号", group_col="批次",
        data_cols=["Vth", "BVdss", "Rds_on"],
        output_path=OUTPUT_XLSX,
        limit_map={r"Vth": 3.0, r"BVdss": 660, r"Rds_on": 2.0},
        show_limit=True, chart_width=18, chart_height=10, marker_size=6,
    )
    assert OUTPUT_XLSX.exists()

    wb = openpyxl.load_workbook(OUTPUT_XLSX)
    assert wb.sheetnames == ["统计汇总", "Vth", "BVdss", "Rds_on"]

    ws = wb["统计汇总"]
    assert ws.cell(1, 1).value == "数据来源"
    stat_header_row = 23
    assert ws.cell(stat_header_row, 1).value == "测试项"
    assert ws.cell(stat_header_row, 15).value == "Limit"
    assert ws.cell(stat_header_row, 16).value == "limit处 CDF"
    assert ws.cell(22, 1).value == "统计汇总"
    assert ws.cell(35, 1).value == "样品编号"

    for sn in ["Vth", "BVdss", "Rds_on"]:
        ws_data = wb[sn]
        ch = ws_data._charts[0]
        assert ch.title is not None
        assert len(ch.series) >= 4, f"{sn}: {len(ch.series)} series"


def test_pipeline_no_limit():
    df = _load_sample()
    process(df=df, id_col="样品编号", group_col="批次", data_cols=["Vth"],
            output_path=OUTPUT_NOLIMIT_XLSX, show_limit=False)
    assert OUTPUT_NOLIMIT_XLSX.exists()


def test_pipeline_defaults():
    df = _load_sample()
    out = FIXTURES / "output_default.xlsx"
    process(df=df, id_col="样品编号", group_col="批次",
            data_cols=["Vth", "BVdss", "Rds_on"], output_path=out,
            chart_title="默认参数", x_label="数据", y_label="CDF")
    assert out.exists()


# ── 多尺度组合输出 ──

SCALE_COMBOS = [
    ("linear", "linear"), ("log", "linear"),
    ("linear", "log"), ("log", "log"),
]


@pytest.mark.parametrize("x_scale,y_scale", SCALE_COMBOS)
def test_scale_combos(x_scale, y_scale):
    df = _load_sample()
    out = FIXTURES / f"output_x{x_scale}_y{y_scale}.xlsx"
    process(df=df, id_col="样品编号", group_col="批次",
            data_cols=["Vth", "BVdss", "Rds_on"], output_path=out,
            limit_map={r"Vth": 3.0, r"BVdss": 660, r"Rds_on": 2.0},
            x_scale=x_scale, y_scale=y_scale, show_limit=True,
            chart_width=18, chart_height=10)
    assert out.exists()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
