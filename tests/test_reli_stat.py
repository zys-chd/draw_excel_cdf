"""reli_stat 单元测试"""

import sys
from pathlib import Path

import pandas as pd
import pytest

# 确保可以 import 项目根目录的 reli_stat 模块
sys.path.insert(0, str(Path(__file__).parent.parent))

from reli_stat import (
    compute_statistics,
    match_limit,
    empirical_cdf,
    median_rank,
    weibull_transform,
    build_summary,
)


# ── empirical_cdf ──────────────────────────────


def test_empirical_cdf_single_group():
    """单组数据，验证 CDF 范围 (0, 1] 且单调递增。"""
    values = pd.Series([2.0, 2.5, 2.3, 2.8, 2.1])
    cdf = empirical_cdf(values)

    assert len(cdf) == 5
    assert (cdf > 0).all()
    assert (cdf <= 1).all()
    # 排序后 CDF 应单调
    sorted_cdf = cdf.loc[values.sort_values().index]
    assert sorted_cdf.is_monotonic_increasing


def test_empirical_cdf_ties():
    """相同值应获得相同 CDF（average rank）。"""
    values = pd.Series([1.0, 2.0, 2.0, 3.0])
    cdf = empirical_cdf(values)

    # 两个 2.0 有相同 rank → 相同 CDF
    assert cdf.iloc[1] == cdf.iloc[2]


def test_empirical_cdf_single_value():
    """单值：rank=1, n=1 → CDF=1.0"""
    values = pd.Series([5.0])
    cdf = empirical_cdf(values)
    assert cdf.iloc[0] == 1.0


def test_empirical_cdf_max_n():
    """n=3: 最大 rank=3 → CDF=3/3=1.0"""
    values = pd.Series([1.0, 2.0, 3.0])
    cdf = empirical_cdf(values)
    assert cdf.max() == 1.0


# ── median_rank ────────────────────────────────


def test_median_rank_basic():
    """验证 median rank 在 (0, 1) 范围内。"""
    values = pd.Series([2.1, 2.5, 2.3, 2.8, 2.0])
    mr = median_rank(values)

    assert len(mr) == 5
    assert (mr > 0).all()
    assert (mr < 1).all()


def test_median_rank_single():
    """单值：MR = (1 - 0.3) / (1 + 0.4) = 0.5"""
    mr = median_rank(pd.Series([5.0]))
    assert abs(mr.iloc[0] - 0.5) < 1e-9


# ── weibull_transform ──────────────────────────


def test_weibull_basic():
    """验证 Weibull 变换单调递增。"""
    mr = pd.Series([0.1, 0.3, 0.5, 0.7, 0.9])
    w = weibull_transform(mr)
    assert w.is_monotonic_increasing
    assert not w.isna().any()


def test_weibull_boundary():
    """MR=0 和 MR=1 → NaN"""
    mr = pd.Series([0.0, 0.5, 1.0])
    w = weibull_transform(mr)
    assert pd.isna(w.iloc[0])
    assert pd.isna(w.iloc[2])
    assert not pd.isna(w.iloc[1])


# ── match_limit ────────────────────────────────


def test_match_limit():
    ids = pd.Series(["A1-001", "A2-002", "B1-003", "C1-004"])
    limit_map = {r"A\d+": 100.0, r"B1": 200.0}

    result = match_limit(ids, limit_map)

    assert result.iloc[0] == 100.0
    assert result.iloc[1] == 100.0
    assert result.iloc[2] == 200.0
    assert pd.isna(result.iloc[3])


def test_match_limit_first_match():
    ids = pd.Series(["AB-001"])
    limit_map = {r"AB": 100.0, r"A": 200.0}
    result = match_limit(ids, limit_map)
    assert result.iloc[0] == 100.0


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
        limit_map={r"G1": 3.5},
    )

    assert list(result.columns) == [
        "id", "group", "variable", "data", "CDF", "weibull", "limit",
    ]
    assert len(result) == 4

    # G1 的 CDF: rank=1,n=2→0.5, rank=2,n=2→1.0
    g1 = result[result["group"] == "A"].sort_values("data")
    assert g1["CDF"].iloc[0] == 0.5
    assert g1["CDF"].iloc[1] == 1.0

    # G1 的 limit = 3.5
    assert (g1["limit"] == 3.5).all()


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
