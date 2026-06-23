# draw_excel_cdf — 可靠性统计 Excel 工具

> 单文件架构。核心绘图函数独立可复用。读取 Excel 测试数据 →
> 按组计算 CDF/Weibull → 输出带统计汇总 sheet + 散点图的 Excel。

---

## 文件结构

```
draw_excel_cdf/
├── draw_excel_cdf.py          # 唯一源码文件
├── tests/
│   └── test_draw_excel_cdf.py
├── DESIGN.md
├── README.md
├── requirements.txt
├── pyproject.toml
└── .gitignore
```

---

## 核心 API

### 一站式入口

```python
from draw_excel_cdf import process

process(
    input_path="data.xlsx",
    id_col="样品编号",
    group_col="批次",
    data_cols=["Vth", "BVdss"],
    output_path="result.xlsx",
    limit_map={r"A组.*": 3.0, r"B组.*": 2.5},
    x_scale="linear",
    y_scale="linear",
    show_limit=True,
    chart_width=20,
    chart_height=12,
    auto_axis=False,
    marker_size=5,
)
```

### 独立绘图函数（可直接复用）

```python
def add_excel_chart(
    wb: Workbook,
    df: pd.DataFrame,         # 已含 CDF/weibull/limit 列的长表
    id_col: str,
    group_col: str,
    data_cols: list[str],
    limit_map: dict | None = None,
    x_axis: str = "data",     # X轴: "data"
    y_axis: str = "CDF",      # Y轴: "CDF" | "weibull"
    x_scale: str = "linear",  # "linear" | "log"
    y_scale: str = "linear",  # "linear" | "log"
    show_limit: bool = True,
    chart_width: float = 20,  # cm
    chart_height: float = 12, # cm
    auto_axis: bool = False,  # True=自动范围, False=手动设置(min/max)
    marker_size: int = 5,
    chart_title: str | None = None,
    x_label: str | None = None,
    y_label: str | None = None,
) -> Workbook:
```

---

## 输出 Excel 结构

```
Sheet 1: "统计汇总"
┌──────┬──────┬───────┬──────┬───────┬───────┬───────┬───────┬──────┬──────┬───────┐
│测试项│总模块│ Group │ 均值 │ 标准差│ 25%分 │ 75%分 │中位数 │ 最小 │ 最大 │变异系数│
│      │  数  │       │      │       │  位   │  位   │       │  值  │  值  │(CV%)  │
├──────┼──────┼───────┼──────┼───────┼───────┼───────┼───────┼──────┼──────┼───────┤
│ Vth  │  15  │   A   │ 2.31 │ 0.12  │ 2.22  │ 2.40  │ 2.30  │ 2.10 │ 2.55 │  5.2  │
│ Vth  │  12  │   B   │ 2.45 │ 0.09  │ 2.39  │ 2.51  │ 2.44  │ 2.32 │ 2.60 │  3.7  │
│ ...  │ ...  │  ...  │ ...  │  ...  │  ...  │  ...  │  ...  │  ... │  ... │  ...  │
└──────┴──────┴───────┴──────┴───────┴───────┴───────┴───────┴──────┴──────┴───────┘

Sheet 2+: 每个数据列一个 sheet（sheet名=数据列名）
┌────────┬───────┬───────┬────────┬──────────┬───────┐
│   id   │ group │  Vth  │  CDF   │  weibull │ limit │
├────────┼───────┼───────┼────────┼──────────┼───────┤
│  ...   │  ...  │  ...  │  ...   │   ...    │  ...  │
└────────┴───────┴───────┴────────┴──────────┴───────┘

  + 散点图（每组不同颜色+marker，无连线）
  + 可选 limit 垂直参考线
```

---

## 散点图特性

- **无连线**：纯散点
- **分组着色**：每组自动分配不同颜色（10色调色板循环）
- **分组 Marker**：每组不同形状（circle, square, diamond, triangle, plus, star, x, dash, dot）
- **Limit 线**：垂直虚线，标注 limit 值
- **坐标轴**：可独立设置 linear/log
- **尺寸**：chart_width/chart_height (cm)
- **范围**：auto_axis=True 自动 / False 手动
- **Marker 尺寸**：marker_size 控制

---

## 算法

- **CDF** = rank / n（简单经验 CDF）
- **Median Rank**: MR = (i - 0.3) / (n + 0.4)
- **limit_map**: `{r"测试项正则": limit_value}` 正则匹配测试项名，首个命中取值
- **Limit**: 按测试项名做正则匹配，同一测试项所有行 limit 一致

---

## 依赖

```
openpyxl>=3.1.0
pandas>=2.0.0
```

---

## 项目结构

```
draw_excel_cdf/
├── draw_excel_cdf.py       # 唯一源码文件 (~950 行)
├── tests/
│   ├── test_draw_excel_cdf.py  # 46 个测试
│   └── fixtures/               # 测试输入/输出
├── DESIGN.md                # 本文件
├── README.md
├── pyproject.toml
├── requirements.txt
├── LICENSE
└── .gitignore
```

### 源码导航 (`draw_excel_cdf.py`)

| 行号范围 | 区域 | 说明 |
|---------|------|------|
| 1~20 | 导入 | openpyxl, pandas, re |
| 25~90 | 常量 | 色板, marker, 表头, 样式 |
| 100~140 | 统计计算 | `empirical_cdf`, `median_rank`, `weibull_transform` |
| 145~165 | limit 匹配 | `get_limit` — 正则匹配测试项名 |
| 170~200 | Weibull 拟合 | `fit_weibull` — 线性回归求 β/η/R² |
| 205~230 | 标签解析 | `_resolve_label` — 正则 map → 字符串 → 默认 |
| 235~260 | 网格线 | `_add_gridlines` |
| 265~340 | **统计汇总** | `build_summary` — 均值/标准差/分位数/Weibull参数 |
| 345~430 | 汇总 sheet 写出 | `write_summary_sheet` — 参数+统计+原始数据 |
| 435~520 | 样式辅助 | `_apply_header_style`, `_apply_data_border` |
| 525~760 | **图表核心** | `add_excel_chart` — 写数据 sheet + 绘制散点图 |
| 765~800 | 文件读取 | `read_file` |
| 805~950 | **流水线入口** | `process` — 加载→计算→写→图→保存 |

---

## 扩展开发指南

### 1. 增加统计指标（新列）

修改位置：
- `SUMMARY_HEADERS`（行 ~75）: 加列名
- `build_summary()`（行 ~265）: 在 `rows.append({...})` 里加计算值
- `write_summary_sheet()` 的 `merge_cells` 要更新列数

```python
# 例：加"峰度"列
SUMMARY_HEADERS = [..., "峰度"]

# build_summary 中:
"峰度": round(values.kurtosis(), 4),
```

### 2. 增加新的 CDF 计算方法

修改 `empirical_cdf()` 或新增函数，然后在 `compute_statistics()`（行 ~235）中替换调用：

```python
sub["CDF"] = sub.groupby(group_col)["data"].transform(your_new_cdf)
```

### 3. 增加新的分布拟合

仿照 `fit_weibull()` 新增函数，返回 `(参数1, 参数2, R²)` 或 `None`。然后在 `build_summary()` 中调用并加入汇总表。

### 4. 增加图表类型或样式

修改 `add_excel_chart()`（行 ~525）:
- 散点图参数在 `_add_gridlines()` 和 marker 样式段
- 如果需要折线图，改 `ScatterChart` 和 `series.graphicalProperties.line`

### 5. 增加数据源类型

`process()` 第一行（行 ~855）判断 `isinstance(data, (str, Path))`：
```python
if isinstance(data, (str, Path)):
    # 文件路径
elif isinstance(data, pd.DataFrame):
    # DataFrame
# 新增：elif hasattr(data, 'read'):   # file-like
```

### 6. 测什么

每次改完跑 `pytest tests/ -v`。新增功能加到 `test_draw_excel_cdf.py`，参考现有测试的模式。

