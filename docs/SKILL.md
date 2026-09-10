# Excel‑Analyzer Skill

> **底层执行入口**：`src/excel_analyzer/cli.py`  
> **通信模式**：子进程命令行调用；业务 JSON 输出到 `stdout`；日志与调试信息输出到 `stderr`；持久化日志输出至 `logs/excel_analyzer.log`  
> **用途**：供智能体理解、调用 Excel 数据分析能力；**本文件不包含项目安装部署步骤，部署请查阅项目根目录 `README.md`**

## 1 使用场景
该Skill用于智能体对本地磁盘Excel文件执行数据分析。仅做文件读取，不会修改原始Excel。
支持5种分析意图：

overview：数据概览，统计行列、字段类型、缺失值、少量样本

stats：数值列描述性统计（均值、中位数、标准差、极值）

sort_filter：排序与条件过滤

trend：时序趋势分析，需要时间列+数值列

anomaly：异常检测，识别重复行、高缺失字段、IQR离群点

## 2 调用参数定义
### 2.1 必选参数
| 参数 | 类型 | 说明 |
|------|------|------|
| `--file` | str | 本地 Excel 文件路径；仅支持后缀 `.xlsx` / `.xls` |
| `--intent` | str | 分析意图，取值见"1 使用场景"章节 |

### 2.2 公共可选参数
| 参数 | 类型 | 说明 |
|------|------|------|
| `--sheet` | int | 工作表索引（从0开始）；默认 0 |
| `--output_json` | str | 可选；将结果JSON写入指定文件；不设置则打印JSON到stdout |

### 2.3 sort_filter 专属可选参数
| 参数 | 类型 | 说明 |
|------|------|------|
| `--sort_col` | str | 待排序列名称 |
| `--sort_asc` | flag | 存在代表升序；不提供默认降序 |
| `--filter_condition` | str | 过滤条件表达式，例：销售额>500 |

### 2.4 trend 专属必传参数
> **注意**：`intent=trend` 时二者**必须同时提供**，否则返回 `MISSING_ARG` 错误。

| 参数 | 类型 | 说明 |
|------|------|------|
| `--time_col` | str | 时间字段列名 |
| `--value_col` | str | 待分析数值列名 |

## 3 返回JSON结构
```json
{
  "success": true,
  "error_code": "",
  "error_msg": "",
  "meta": {
    "file_path": "",
    "sheet_name": "",
    "intent": ""
  },
  "result": {}
}
```
- `success=true`：业务结果放置于`result`；datetime 时间对象自动序列化为 ISO‑8601 字符串。
- `success=false`：业务判断**只使用 `error_code` 字段**，禁止依赖 `error_msg` 文本匹配。

完整错误码表：

| error_code | 含义 |
|------------|------|
| `FILE_NOT_FOUND` | 文件路径不存在 |
| `INVALID_SUFFIX` | 文件后缀不是 `.xlsx` / `.xls` |
| `SHEET_NOT_EXIST` | 指定工作表不存在 |
| `INVALID_INTENT` | `intent` 不在合法集合 |
| `MISSING_ARG` | 当前意图缺少必填参数 |
| `DATA_TYPE_ERR` | 数据解析失败（如 `trend` 时间列无法解析为日期） |
| `INTERNAL_ERROR` | 工具内部未知异常 |
| `NEED_USER_CONFIRM` | ✨需要用户确认是否添加目录进入信任白名单，result携带confirm_info |

## 4 调用示例
前提：运行环境已配置好 PYTHONPATH，指向项目根目录

示例 1：获取数据概览

```bash
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent overview --sheet "销售数据"
```

示例 2：时序趋势分析

```bash
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent trend --sheet "销售数据" --time_col "日期" --value_col "销售额"
```
示例 3：排序筛选

```bash
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent sort_filter --sheet "销售数据" --sort_col "销售额" --sort_asc --filter_condition "销售额>500"
```

## 5 Agent 调用规范（调用方必须遵守）
使用子进程 subprocess 调用 cli 脚本；必须设置环境变量PYTHONPATH为项目根目录。

仅解析 stdout 作为业务返回结果；stderr 只用于调试，不得送入大模型上下文。

本工具只能访问本地磁盘文件，不支持网络远程文件。

本项目在 `docs/` 目录提供适配 Nanobot 的 `skill.py` 样例文件，与 `SKILL.md` 位于同一目录。该样例仅面向 Nanobot 框架；对接其他 Agent 框架需要自行改写适配。

使用部署：将 `SKILL.md`、`skill.py` 两份文件复制到 Nanobot 的自建技能目录完成技能注册；**禁止放入 ai_excel_agent/src 源码目录**。

运行说明：样例优先自动探测ai_excel_agent项目根目录；探测失败时设置环境变量 `AGENT_PROJECT_ROOT` 指定项目根路径。`agent_apply_confirm`来自`src.excel_analyzer.whitelist_manager`内部Python API，无CLI命令。

### Nanobot 集成特殊说明
当 `excel_analyze` 返回 `error_code=="NEED_USER_CONFIRM"`，业务层需要把返回的 `error_msg` 展示给终端用户；获取用户回复文本后，调用 skill 内部辅助函数 `excel_handle_user_confirm(user_input, session_ctx)` 完成确认与重跑分析。
`excel_handle_user_confirm` 不属于 `SKILL_TOOLS` 工具列表，**不暴露给 function‑call**，仅供 Nanobot 上层业务代码直接调用。

## 6 安全契约（调用 Agent 强制实现）
> ⚠️ **CLI 层内置持久化白名单与用户确认机制**

首次生成配置文件时，自动将项目内置 tests/test_data 加入默认信任白名单；若配置已存在，不会修改用户已有白名单。

TTY 终端交互模式：CLI 内部直接交互式询问用户确认；

被 subprocess 子进程调用（Agent/Skill）：返回错误码NEED_USER_CONFIRM，result.confirm_info包含候选目录，Agent 必须向人类展示路径，拿到显式确认，再调用白名单添加接口，之后重新发起分析请求。

最小权限：防范路径遍历攻击，除首次初始化内置的tests/test_data外，其余所有目录添加必须经过用户显式确认，程序不会自动新增其他目录。

用户确认机制：当待读取文件不在信任白名单内，Agent 必须向用户展示完整文件路径，获取用户显式确认之后才允许访问；白名单内可跳过确认

敏感信息保护：工具返回结果会附带少量单元格样本；Agent 需要识别手机号、身份证等敏感字段并做脱敏掩码，再提交给大模型。

日志可追溯：工具自身日志不会保存完整 Excel 原始内容，仅记录文件路径、意图、错误码；Agent 不可将密钥、API 密钥传入工具参数。

agent_apply_confirm仅为 Python 内部 API，没有对应 CLI 命令行参数，只能由上层 Agent/Python 代码调用，不能通过命令行参数调用。

收到NEED_USER_CONFIRM后不可直接重试 cli；需先向用户确认，调用agent_apply_confirm完成目录添加，再重新调用分析接口。

## 7 低 Token 与性能优化契约
工具已经实现的优化，Agent 调用时不需要再做重复处理：

业务计算下沉 Python 脚本，仅返回统计指标 + 少量样本，不输出完整表格，减少 LLM 输入 token。

overview、sort_filter 仅返回头部少量样本行。

trend 参数校验前置到文件 IO 读取之前，参数错误直接返回，避免无效读取大 Excel。

## 8 独立测试方式
在项目根目录执行单元测试，验证工具本身功能正确性：

```bash
python -m unittest tests.test_basic -v
```

测试用例共 10 个，覆盖异常输入、边界脏数据、正常业务通路。