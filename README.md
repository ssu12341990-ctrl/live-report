# live-report

## ADQ 输入与生成流程

1. 在 `data/adq_input_template.xlsx` 的 `ADQ输入模板` sheet 中填入数据（从第5行开始追加每个投流组）。
2. `必填` 字段不能为空；`选填` 字段（关注人数/支付金额/退款金额/备注）可留空。
3. 运行以下脚本生成报表：
   - `python generate_filled_adq.py`（基础版，输出 `ADQ直播数据统计_已填数据.xlsx`）
   - `python generate_advanced_adq_report.py`（高级版，输出 `ADQ高级投流分析_已填数据.xlsx`）
4. 两个脚本均使用同一个输入源：`data/adq_input_template.xlsx`。