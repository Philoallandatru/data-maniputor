# PandasAI + 本地 LLM Excel 分析助手

一个基于 `Streamlit + pandas + PandasAI + 本地 OpenAI-compatible API` 的 Excel / CSV 数据分析 MVP。

适合的核心流程：

```text
上传文件 -> 预览数据 -> 配置本地模型 -> 聊天提问 -> 返回文本 / 表格 / 图表
```

默认分析方向偏向 SSD 性能测试数据，例如：

- 多个固件版本对比
- workload 性能退化识别
- IOPS / 带宽下降分析
- latency / P95 / P99 上升分析
- 输出中文结论摘要

## 运行前提

- 推荐 Python `3.11`
- 当前 `pandasai-litellm` 通常要求 Python `<3.12`
- 如果你使用 Python `3.12+`，很可能会出现安装或导入失败

## 安装方式

```bash
pip install -r requirements.txt
```

## 启动方式

```bash
streamlit run app.py
```

启动后打开浏览器中的 Streamlit 页面即可使用。

## 支持的文件格式

- `.csv`
- `.xlsx`
- `.xlsm`
- `.xls`

说明：

- `.xls` 读取依赖 `xlrd`
- 默认最多读取前 `50,000` 行
- 可以在侧边栏调整最大读取行数
- 预览区默认只展示前 `100` 行

## 界面说明

侧边栏包含：

- `API Base URL`
- `模型名`
- `API Key`
- `temperature`
- `max_tokens`
- `最大读取行数`
- `是否保存图表`

主界面包含：

- 文件上传
- Excel sheet 选择
- 数据概览
- 字段信息
- 数据预览
- 聊天分析窗口

## LM Studio 配置

示例：

```text
API Base: http://localhost:1234/v1
API Key: lm-studio
Model: /v1/models 中返回的模型名
```

说明：

- 应用会自动去掉 API Base 末尾多余的 `/`
- 如果模型名没有 `openai/` 前缀，应用会自动补成 `openai/<你的模型名>`

## Ollama 配置

如果你启用了 Ollama 的 OpenAI-compatible 接口，可参考：

```text
API Base: http://localhost:11434/v1
API Key: ollama
Model: qwen2.5:32b
```

应用会自动把模型名标准化为：

```text
openai/qwen2.5:32b
```

## vLLM 配置

示例：

```text
API Base: http://localhost:30001/v1
API Key: EMPTY
Model: qwen-local
```

## 常见问题

### 1. `pandasai-litellm` 安装失败

优先检查 Python 版本：

```text
python --version
```

如果是 Python `3.12+`，请切换到 Python `3.11` 后重新安装。

### 2. 上传 `.xls` 失败

请确认 `xlrd` 已安装：

```bash
pip install xlrd
```

### 3. 提示无法连接模型服务

请检查：

1. 本地模型服务是否已经启动
2. API Base 是否包含 `/v1`
3. 模型名是否和 `/v1/models` 返回一致
4. API Key 是否为非空字符串

### 4. 模型返回结果异常或分析失败

请检查：

1. 当前模型是否具备代码生成能力
2. PandasAI 与 `pandasai-litellm` 是否安装成功
3. 当前数据列名是否过于混乱或缺少关键字段

### 5. 图表没有正常显示

请检查：

1. 本地模型是否正确生成了绘图代码
2. `exports/charts` 目录是否可写
3. 返回的图表路径是否存在

## 安全说明

本项目默认只连接你在页面中配置的本地 OpenAI-compatible API，不会主动调用 OpenAI 官方云端接口。

但需要注意：

- PandasAI 会让 LLM 生成 Python 代码并执行
- 这意味着它并不是纯文本问答
- 请仅在可信的本地环境中使用

建议：

- 仅分析可信文件
- 仅连接可信的本地模型服务
- 如果后续需要多人共享部署，建议增加 Docker 沙箱

## 适合的场景

- 数万行以内 Excel / CSV 的快速探索分析
- SSD 性能测试结果对比
- 固件版本回归初筛
- 本地离线或半离线分析
- 给测试负责人快速生成中文摘要

## 不适合的场景

- 百万行以上超大数据集
- 严格生产级沙箱隔离要求
- 复杂 BI 报表平台替代
- 多用户权限与审计系统

如果后续需要更大数据规模，建议升级为：

```text
DuckDB / Parquet / 固定分析规则 / 报告导出 / Docker 沙箱
```

# data-maniputor
