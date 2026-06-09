# toolboxscript

一个可持续扩展的脚本合集仓库。

## 现在有啥

- `riskbird-check`：风鸟企业核验

## 运行方式

```bash
export RISKBIRD_COOKIE='...'
toolboxscript riskbird-check --input /path/to/企业数据-整理-Bad.xlsx
```

也可以直接用：

```bash
python -m toolboxscript riskbird-check --input /path/to/企业数据-整理-Bad.xlsx
```

## 每次运行会留下些什么

每次都会生成一个独立目录：

```text
runs/YYYY-MM-DD_HHMMSS_riskbird-check/
```

目录里会有：

- `log.jsonl`：阶段日志、逐条查询日志、重试日志
- `summary.json`：本次运行汇总
- `raw/`：搜索和详情接口原始响应
- `output/`：最终导出的 xlsx/csv
- `input.xlsx`：输入文件快照（如果开启复制）

## 风控

- 默认单线程、低频率查询
- 每条企业之间有随机间隔
- 失败会退避重试
- 搜索命中但证据不足时，默认进 `需复核`
- 只有多种检索都没有结果时才算 `未查到`

## 扩展新脚本

后续新增脚本时，优先复用：

- `toolboxscript.common`：日志、重试、路径、文本处理、Excel
- `toolboxscript.connectors`：外部服务适配器
- `toolboxscript.runners`：脚本业务流

只需要再加一个 runner 和 CLI 子命令即可。

