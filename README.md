# pwkkx — 10kV 配电线路供电可靠性计算

基于主线与分支分段数据的供电可靠性指标计算（SAIDI、SAIFI、ASAI），支持电缆/架空敷设方式权重与参数配置。

## 快速开始

```bash
# 依赖
pip install pandas openpyxl

# 指定输入 Excel，输出使用默认目录
python main.py -i document/10kV安54新窑线.xlsx

# 指定输入与输出
python main.py -i document/10kV景704景水线.xlsx -o workspace/result/景水线_结果.xlsx

# 使用自定义参数文件
python main.py -i <输入.xlsx> -o <输出.xlsx> -c config/reliability_params.json
```

## 项目结构

```
pwkkx/
├── main.py                 # 主入口（泛化框架，-i / -o / -c）
├── config/
│   └── reliability_params.json   # 常量、Sheet 名、字段映射
├── document/
│   ├── 10kV配电线路供电可靠性计算算法逻辑说明书.md   # 算法逻辑说明
│   ├── 技术方案与框架算法对照说明.md
│   ├── reliability_algorithm.py
│   └── reliability_calculation.py
└── workspace/
    ├── reliability_framework.py  # 与 main 同逻辑，可单独运行
    ├── reliability_calc_jingshuixian.py
    └── result/                    # 默认输出目录
```

## 输入输出

- **输入**：含「主线」「分支」两个 Sheet 的 Excel，列名通过 `config/reliability_params.json` 的 `field_mappings` 映射。
- **输出**：含「主线分段明细」「分支分段明细」「指标汇总」三个 Sheet 的 Excel；未指定 `-o` 时写入 `workspace/result/<输入文件名>_可靠性计算结果.xlsx`。

## 算法与文档

- 算法整体逻辑、公式、常量与流程见：[document/10kV配电线路供电可靠性计算算法逻辑说明书.md](document/10kV配电线路供电可靠性计算算法逻辑说明书.md)
- 与技术方案的对照见：[document/技术方案与框架算法对照说明.md](document/技术方案与框架算法对照说明.md)

## 许可证

MIT
