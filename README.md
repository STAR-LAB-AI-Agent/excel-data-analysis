# AI‑Excel‑Agent

课程实验选题：01 AI‑Excel 数据分析  
实验模式：Script/CLI + Skill，实现智能体调用 Excel 数据分析能力。

---

## 项目简介

基于 `pandas`、`openpyxl` 将 Excel 分析能力封装为独立命令行 CLI 程序。智能体通过子进程调用本工具，完成 Excel 的数据概览、统计、排序筛选、时序趋势、异常检测。

**特点**：
- 工具仅执行文件读取，**不会修改原始 Excel 文件**
- 智能体调用能力说明文档：`docs/SKILL.md`
- 支持 5 种分析意图：overview / stats / sort_filter / trend / anomaly

---

## 用户场景

本工具适用于以下场景：

1. **数据快速概览**：用户拿到一份 Excel 表格，想快速了解有多少行、多少列、有哪些字段、是否存在缺失值。
2. **数值统计分析**：用户需要对表格中的数值列进行描述性统计，了解均值、中位数、标准差、极值等。
3. **数据筛选与排序**：用户需要按特定条件筛选数据（如销售额 > 500），并按某列排序查看。
4. **时序趋势分析**：用户有时间列和数值列，想了解指标随时间的变化趋势（上升/下降/平稳）。
5. **异常检测**：用户需要识别表格中的重复行、缺失率过高的列、以及数值列中的离群点。

---

## 开源第三方依赖（任务书要求）

| 库名称 | 用途 | 许可证 |
|--------|------|--------|
| `pandas` | 表格解析、统计、时序数据处理 | BSD‑3‑Clause |
| `openpyxl` | `.xlsx` / `.xls` 文件解析引擎 | MIT |

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
│   └── cli.py                   # CLI 对外入口
├── config/                      # ✨运行时配置目录，git忽略，不提交仓库
│   └── excel_agent_config.json  # 白名单持久化配置，首次运行自动生成
├── docs/
│   └── SKILL.md                 # 给智能体使用的 Skill 调用文档
│   └── skill.py                 # 适配 Nanobot 的 skill.py 样例
├── tests/
│   ├── test_basic.py            # 单元测试，共 10 个用例
│   ├── generate_test_data.py    # 生成测试样本 Excel
│   ├── debug_cli.py             # 调试脚本
│   └── test_data/               # 测试数据文件
│       └── sample_test.xlsx
├── logs/                        # 运行自动生成，持久化日志
├── requirements.txt
└── .gitignore
```

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
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent overview --sheet "销售数据"
```
## Agent 接入简要说明
- 智能体以子进程 `subprocess` 调用 `src/excel_analyzer/cli.py`
- 完整参数、返回格式、安全契约请阅读 `docs/SKILL.md`
- `stdout` 为业务 JSON 输出；`stderr` 仅用于调试日志，**禁止送入大模型**
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

### 2. 执行全套单元测试（10 条用例）

```bash
python -m unittest tests.test_basic -v
```

覆盖异常场景、边界脏数据、正常业务功能，满足任务书 5‑10 条测试用例要求。

## 安全说明
- 工具仅读取 Excel，**不会修改原文件**；但可读取本机任意可读 Excel
- CLI层已内置持久化白名单与用户确认机制；区分终端交互模式 / Agent非交互模式。Agent收到 `NEED_USER_CONFIRM` 错误码时，需要向用户展示完整文件路径，拿到用户显式确认后再执行后续操作。
- 日志采用轮转文件日志，不会保存完整 Excel 原始内容，规避敏感数据泄露
- **禁止将 API Key、密钥提交到代码仓库**

## 低 Token 优化说明
- 统计、过滤、异常计算下沉 Python 脚本，仅返回统计指标 + 少量样本，减少送入 LLM 的 Token 数量
- 参数校验前置，在磁盘 IO 读取文件之前拦截非法参数，避免大文件无效 IO
- 结果只返回有限样本行数，不返回完整表格

## 已知限制
- 仅支持本地磁盘 Excel，不支持网络远程文件
- 可选扩展功能（未实现）：自动生成分析图表

## 许可证
课程实验项目代码使用 MIT 许可证；第三方库遵循其各自开源许可证。