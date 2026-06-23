# draw_excel_cdf

可靠性数据分析工具 — 读取测试数据，按组计算 CDF / Weibull 变换，输出带统计汇总和散点图的 Excel 文件。

## 安装

```bash
pip install -r requirements.txt
```

依赖: `openpyxl>=3.1`, `pandas>=2.0`

## 快速开始

```python
from draw_excel_cdf import process

# 传入 DataFrame
process(
    data=df,                          # DataFrame 或 Excel 路径
    id_col="样品编号",
    group_col="批次",
    data_cols=["Vth", "BVdss", "Rds_on"],
    output_path="result.xlsx",
    limit_map={"Vth": 3.0, "BVdss": 650},
)
```

## 输出 Excel 结构

| Sheet | 内容 |
|-------|------|
| 统计汇总 | 参数信息 + 分组统计（均值/标准差/分位数/Weibull βηR²/limit处CDF） + 原始数据 |
| 各数据列 | id / group / 数据 / CDF / weibull / limit + **散点图** |

## 图表特性

- 按组着色散点图，不同 marker 形状
- X/Y 轴独立 linear/log 缩放
- 可自定义轴范围、图表尺寸、marker 大小
- limit 参考线（正则匹配测试项名）
- 图例右侧竖排，绘图区黑边框，浅灰网格线

## License

MIT
