# draw_excel_cdf

可靠性统计 Excel 工具——读取测试数据，按组计算 CDF（Median Rank）和 Weibull 变换，
输出带统计汇总和散点图的 Excel 文件。

## 安装

```bash
pip install -r requirements.txt
```

## 使用

```python
from draw_excel_cdf import process

process(
    input_path="data.xlsx",
    id_col="样品编号",
    group_col="批次",
    data_cols=["Vth", "BVdss"],
    output_path="output.xlsx",
    limit_map={r"A组.*": 3.0},
    x_scale="linear",
    y_scale="log",
)
```

## 独立图表函数

```python
from draw_excel_cdf import add_excel_chart
from openpyxl import Workbook

wb = Workbook()
# ... 填充数据 ...
wb = add_excel_chart(wb, df, "id", "group", ["Vth"], show_limit=True)
```

## 依赖

- Python >= 3.9
- openpyxl >= 3.1.0
- pandas >= 2.0.0
