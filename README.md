# AI‑Excel‑Agent

课程实验选题：01 AI‑Excel 数据分析  
实验模式：Script/CLI + Skill，实现智能体调用 Excel 数据分析能力。

---

## 项目简介

基于 `pandas`、`openpyxl` 将 Excel 分析能力封装为独立命令行 CLI 程序。智能体通过子进程调用本工具，完成 Excel 的数据概览、统计、排序筛选、时序趋势、异常检测、图表生成，并可**一次性产出完整分析报告**。

**特点**：
- 工具仅执行文件读取，**不会修改原始 Excel 文件**
- 智能体调用能力说明文档：`docs/SKILL.md`
- 支持 7 种分析意图：overview / stats / sort_filter / trend / anomaly / chart / **report**
- `report` 为**推荐入口**：单次调用跑完六个分段，直接返回可交付的 Markdown 报告 + 图表路径 + 交付自查表
- `report` 内置**离群点校验**：自动识别“表面暴跌/暴涨其实由离群点驱动”的伪趋势

---

## 用户场景

本工具适用于以下场景：

1. **数据快速概览**：用户拿到一份 Excel 表格，想快速了解有多少行、多少列、有哪些字段、是否存在缺失值。
2. **数值统计分析**：用户需要对表格中的数值列进行描述性统计，了解均值、中位数、标准差、极值等。
3. **数据筛选与排序**：用户需要按特定条件筛选数据（如销售额 > 500），并按某列排序查看。
4. **时序趋势分析**：用户有时间列和数值列，想了解指标随时间的变化趋势（上升/下降/平稳）。
5. **异常检测**：用户需要识别表格中的重复行、缺失率过高的列、以及数值列中的离群点。
6. **图表生成（可选功能）**：用户需要将分析结果可视化，生成柱状图、折线图、箱线图或直方图，输出为 PNG 文件。
7. **一次性完整报告（推荐）**：用户一次提出多个分析需求（如“概览+统计+筛选+趋势+异常检测+柱状图”），
   用 `intent=report` 单次调用即可得到一份含全部关键数值与图表的完整报告，不需逐项跑、也不会漏答其中某一问。

---

## 开源第三方依赖（任务书要求）

| 库名称 | 用途 | 许可证 |
|--------|------|--------|
| `pandas` | 表格解析、统计、时序数据处理 | BSD‑3‑Clause |
| `openpyxl` | `.xlsx` / `.xls` 文件解析引擎 | MIT |
| `matplotlib` | 图表生成（可选功能） | PSF / BSD-compatible |

依赖版本记录于 `requirements.txt`。

---

## 目录结构

```text
ai_excel_agent/
├── src/excel_analyzer/          # 核心业务源码
│   ├── __init__.py
│   ├── schemas.py               # 常量、返回结构定义
│   ├── exceptions.py            # 业务自定义异常
│   ├── logger_cfg.py            # 轮转日志配置
│   ├── analyzer_core.py         # Excel 分析业务逻辑
│   ├── data_cleaner.py          # 数据清洗工具
│   ├── chart_generator.py       # 图表生成（可选功能）
│   ├── report_builder.py        # ✨一次性完整报告编排与 Markdown 渲染（intent=report）
│   ├── whitelist_manager.py     # 白名单持久化与校验
│   └── cli.py                   # CLI 对外入口
├── config/                      # ✨运行时配置目录，git忽略，不提交仓库
│   └── excel_agent_config.json  # 白名单持久化配置，首次运行自动生成
├── docs/
│   ├── SKILL.md                 # 给智能体使用的 Skill 调用文档
│   └── skill.py                 # 适配 Nanobot 的 skill.py 样例
├── tests/
│   ├── test_basic.py            # 基础测试，11 个用例
│   ├── test_advanced.py         # 进阶测试，12 个用例
│   ├── test_report.py           # ✨一次性完整交付契约测试，12 个用例
│   ├── test_encoding.py         # 输出编码契约测试（stdout 恒为 UTF-8），5 个用例
│   ├── test_doc_consistency.py  # 文档一致性契约测试，6 个用例
│   ├── generate_test_data.py    # 生成测试样本 Excel
│   ├── debug_cli.py             # 调试脚本
│   └── test_data/               # 测试数据文件
│       └── sample_test.xlsx
├── check_deps.py                # 查询依赖版本，用于生成 requirements.txt
├── logs/                        # 运行自动生成，持久化日志（git忽略）
├── outputs/                     # 运行时生成，图表输出（git忽略）
├── requirements.txt
└── .gitignore
```

> 项目根目录**不应**残留分析产物（如 `*.json` 结果转储、`analysis_out/` 等）；
> 需要保存结果时请使用 `--output_json` 指定到仓库外或已忽略的目录。

---

## 环境部署
### 1. 创建虚拟环境（conda 示例）
```bash
conda create -n models python=3.11
conda activate models
```
### 2. 安装依赖
在项目根目录执行：

```bash
pip install -r requirements.txt
```
### 3. 运行前置条件

调用 CLI 脚本必须设置环境变量 `PYTHONPATH` 指向项目根目录：

```powershell
# Windows PowerShell
$env:PYTHONPATH = "$PWD"
```

```bash
# Linux / macOS
export PYTHONPATH="$PWD"
```
### 4. 手动 CLI 调用示例

```bash
# 【推荐】一次性完整报告：stdout 直接就是最终报告正文（Markdown）
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent report --format md

# 单个意图示例
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent overview --sheet "销售数据"
```

#### `intent=report` 常用参数

| 参数 | 说明 |
|------|------|
| `--format json/md` | `md`：stdout 只输出报告正文；`json`（默认）：输出完整 JSON |
| `--sort_col` / `--sort_asc` | 报告内排序段；report 下 `--sort_asc` 缺省为升序 |
| `--filter_condition` | 报告内过滤条件，例：`销售额>500` |
| `--time_col` / `--value_col` | 报告内趋势段列名（不传自动探测） |
| `--x_col` / `--y_col` / `--chart_type` / `--title` | 报告内图表设置（不传自动探测） |

## Agent 接入简要说明
- 智能体以子进程 `subprocess` 调用 `src/excel_analyzer/cli.py`
- **多意图请求优先使用 `--intent report --format md`**：一次调用即得到完整报告，避免多次调用与自行拼装导致的交付不完整
- 完整参数、返回格式、安全契约请阅读 `docs/SKILL.md`
- `stdout` 为业务 JSON 输出；`stderr` 仅用于调试日志，**禁止送入大模型**（`--format md` 时 stdout 为报告正文）
- **交付契约**：`report` 返回的 `markdown` 全文与 `attachments` 中的图片**必须在同一条消息中发出**，禁止只发图表、禁止拆成两条气泡
- Agent 收到NEED_USER_CONFIRM禁止循环重试 cli，必须完成用户确认、调用agent_apply_confirm写入配置后，再重新发起分析；
- agent_apply_confirm是 Python 内部 API，无 CLI 命令，需要直接导入模块调用。
- 项目在docs/目录提供仅适配 Nanobot 的 skill.py 样例，和SKILL.md同目录；对接其他 Agent 框架需要自行修改适配。使用时将 SKILL.md、skill.py 复制至 Nanobot 的自建技能目录完成技能注册，不可放入本项目 src 源码目录,并将ai_excel_agent技能文件夹置于Agent可操作文件夹下。
- skill.py 会自动探测 ai_excel_agent 项目根目录；探测失败，设置环境变量AGENT_PROJECT_ROOT指向项目根目录。

### Python 调用最小示例

```python
import subprocess
import json
import os

root = os.path.abspath(".")
env = os.environ.copy()
env["PYTHONPATH"] = root

cmd = [
    "python",
    os.path.join(root, "src/excel_analyzer/cli.py"),
    "--file", r"tests/test_data/sample_test.xlsx",
    "--intent", "overview"
]

proc = subprocess.run(cmd, capture_output=True, text=True, env=env)
resp = json.loads(proc.stdout)
print(resp)
```

### 一次性完整报告调用示例（推荐）

```python
import subprocess, json, os

root = os.path.abspath(".")
env = os.environ.copy(); env["PYTHONPATH"] = root

cmd = [
    "python", os.path.join(root, "src/excel_analyzer/cli.py"),
    "--file", r"tests/test_data/sample_test.xlsx",
    "--intent", "report",          # 单次调用跑完六个分段
    "--format", "json",            # md 则 stdout 直接是报告正文
    "--sort_col", "销售额", "--sort_asc",
    "--filter_condition", "销售额>500",
]
resp = json.loads(subprocess.run(cmd, capture_output=True, text=True, env=env).stdout)

markdown    = resp["result"]["markdown"]           # 直接作为回复正文
attachments = resp["result"]["attachments"]        # 必须与 markdown 同一条消息发出
checklist   = resp["result"]["delivery_checklist"]  # 发送前逐条自查
```

## 白名单管理
本项目白名单持久化存储于 `config/excel_agent_config.json`。
- **配置文件首次自动生成时，会将项目内置 `tests/test_data` 设置为默认信任白名单；如果配置文件已存在，则不会修改用户已有的白名单。**
- 除初始化默认目录外，新增信任目录均需要用户显式确认，程序不会自动添加目录。

```bash
# 查看当前信任白名单
python src/excel_analyzer/cli.py --list-whitelist

# 删除指定序号的白名单目录
python src/excel_analyzer/cli.py --remove-whitelist-idx 0
```
- TTY 终端交互模式：控制台交互式询问用户确认目录添加；
- Agent 子进程调用模式：访问不在白名单文件返回错误码 NEED_USER_CONFIRM；Agent 需要向用户展示完整路径，获取显式确认后再完成目录添加与分析。

## 测试流程
> ⚠️单元测试前置：删除旧配置 `config/excel_agent_config.json`，让程序自动生成包含`tests/test_data`的默认白名单。单元测试以subprocess非TTY模式运行，不会弹出交互询问。
### 1. 生成测试 Excel 样本

```bash
python tests/generate_test_data.py
```

### 2. 执行全套单元测试

```bash
python -m unittest tests.test_basic -v            # 基础测试（11 条）
python -m unittest tests.test_advanced -v        # 进阶测试（12 条）
python -m unittest tests.test_report -v          # 一次性完整交付契约测试（12 条）
python -m unittest tests.test_encoding -v        # 输出编码契约测试（5 条）
python -m unittest tests.test_doc_consistency -v # 文档一致性契约测试（6 条）
```

也可一次跑完：

```bash
python -m unittest tests.test_basic tests.test_advanced tests.test_report tests.test_encoding tests.test_doc_consistency
```

覆盖异常场景、边界脏数据、正常业务功能、图表可选功能、数据清洗、算法增强、**一次性完整交付契约**、**输出编码契约**、**文档一致性契约**，测试用例共 **46 个**（test_basic 11 + test_advanced 12 + test_report 12 + test_encoding 5 + test_doc_consistency 6）。

其中 `tests.test_report` 锁定以下回归点（即“只发图、其余问题没回答”的故障形态）：
六段标题齐全、anomaly 四要素、chart 带 `output_path` 与分组数值、`attachments` 非空、
`delivery_checklist` 含“同一条消息”要求、`--format md` 输出纯 Markdown。

## 安全说明
- 工具仅读取 Excel，**不会修改原文件**；但可读取本机任意可读 Excel
- CLI层已内置持久化白名单与用户确认机制；区分终端交互模式 / Agent非交互模式。Agent收到 `NEED_USER_CONFIRM` 错误码时，需要向用户展示完整文件路径，拿到用户显式确认后再执行后续操作。
- 日志采用轮转文件日志，不会保存完整 Excel 原始内容，规避敏感数据泄露
- **禁止将 API Key、密钥提交到代码仓库**

## 低 Token 优化说明
- 统计、过滤、异常计算下沉 Python 脚本，仅返回统计指标 + 少量样本，减少送入 LLM 的 Token 数量
- 参数校验前置，在磁盘 IO 读取文件之前拦截非法参数，避免大文件无效 IO
- 结果只返回有限样本行数，不返回完整表格
- `chart` 输出为 PNG 文件，只返回文件路径和大小，**不返回 base64 图像数据**。
- `report` 把六个分段的**拼装下沉到工具内部**：智能体只需一次调用 + 一次转发，
  既省去多轮上下文，也消除“拼装断档导致漏答”的风险；`chart` 段额外给出分组数值表，
  报告不依赖看图即可得出区域对比结论。

## 已知限制
- 仅支持本地磁盘 Excel，不支持网络远程文件
- 图表生成仅支持柱状图、折线图、箱线图、直方图四种类型
- `report` 每次调用只生成一张图表（默认 bar，按分类列取均值）
- 如需其他 Agent 框架集成，需重写适配层（skill.py仅适配 Nanobot）

## 许可证
课程实验项目代码使用 MIT 许可证；第三方库遵循其各自开源许可证。
