# reli-stat — 可靠性统计 Excel 工具

> 单文件架构。核心绘图函数独立可复用。读取 Excel 测试数据 →
> 按组计算 CDF/Weibull → 输出带统计汇总 sheet + 散点图的 Excel。

---

## 文件结构

```
reli-stat/
├── reli_stat.py          # 唯一源码文件
├── tests/
│   └── test_reli_stat.py
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
from reli_stat import process

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
- **limit_map**: `{"测试项名": limit_value}` 每个测试项一个固定 limit 值
- **Limit**: 按测试项名精确匹配，不随样品 ID 变化

---

## 依赖

```
openpyxl>=3.1.0
pandas>=2.0.0
```
