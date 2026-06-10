# riskbird-company-check

风鸟企业经营核验脚本。

## 做什么

读取 Excel 里的 `企业名称` 和 `统一社会信用代码`，只调用风鸟搜索接口，识别搜索结果里的经营状态，并把结果写回表格。

## 输出

每次运行会生成一个目录：

```text
runs/YYYY-MM-DD_HHMMSS_riskbird-company-check/
```

里面会有：

- `logs/log.jsonl`
- `raw/search/*.json`
- `output/企业经营核验结果.xlsx`
- `output/results.json`
- `summary.json`

## 怎么跑

```bash
cd /Users/mac/workspace/active/toolbox-scripts/riskbird-company-check
echo 'RISKBIRD_TOKEN=你的token' > .env
python3 main.py --input /path/to/企业数据-整理-Bad.xlsx
```

也支持：

- `cookie.py`，写 `TOKEN = "..."` 或 `RISKBIRD_TOKEN = "..."`
- `--token-file ./cookie.py`
- 环境变量 `RISKBIRD_TOKEN`

## 规则

- 只查搜索接口，不再查详情页
- 单条企业顺序执行，默认每条间隔 4-8 秒
- 同一企业需要从信用代码兜底到名称查询时，中间也会等待
- 遇到 401、429 或 `operation:too:fast` 会重试；重试后仍受限会停止继续查询
- Excel 会边跑边写入
