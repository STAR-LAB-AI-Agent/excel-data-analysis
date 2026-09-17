# Excel‑Analyzer Skill

> **底层执行入口**：`src/excel_analyzer/cli.py`  
> **通信模式**：子进程命令行调用；业务结果输出到 `stdout`（默认 JSON；`intent=report --format md` 时为 Markdown 报告正文）；日志与调试信息输出到 `stderr`；持久化日志输出至 `logs/excel_analyzer.log`  
> **用途**：供智能体理解、调用 Excel 数据分析能力；**本文件不包含项目安装部署步骤，部署请查阅项目根目录 `README.md`**

## 1 使用场景
该Skill用于智能体对本地磁盘Excel文件执行数据分析。仅做文件读取，不会修改原始Excel。
支持 **7 种**分析意图：

overview：数据概览，统计行列、字段类型、缺失值、少量样本

stats：数值列描述性统计（均值、中位数、标准差、极值、**变异系数、偏度**）

sort_filter：排序与条件过滤

trend：时序趋势分析，需要时间列+数值列；**默认自动清洗无法解析的时间行**

anomaly：异常检测，识别重复行、高缺失字段、IQR离群点、**Z-Score 离群点**

chart：生成分析图表（可选功能），输出 PNG 文件到 outputs/ 目录

**report：✨一次性完整报告（推荐入口）**。单次调用内依次跑完 overview→stats→sort_filter→trend→anomaly→chart，
并直接返回一份**可直接作为回复正文的完整 Markdown 报告**（含全部关键数值、anomaly 四要素、chart 的 output_path），
同时返回 `attachments`（图表路径）与 `delivery_checklist`（交付自查表）。
**当用户请求涉及 2 个以上意图时，必须优先使用 `report`，禁止再逐意图多次调用。**

## 2 调用参数定义
### 2.1 必选参数
| 参数 | 类型 | 说明 |
|------|------|------|
| `--file` | str | 本地 Excel 文件路径；仅支持后缀 `.xlsx` / `.xls` |
| `--intent` | str | 分析意图，取值见"1 使用场景"章节 |

### 2.2 公共可选参数
| 参数 | 类型 | 说明 |
|------|------|------|
| `--sheet` | str | 工作表名称或索引数字（从0开始）；默认 0 |
| `--output_json` | str | 可选；将结果JSON写入指定文件（按 UTF-8）；不设置时：`--format json` 打印到 stdout，`report --format md` 下 stdout 仍恒为 Markdown 正文（见 9.1） |

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
| `--no_clean` | flag | 关闭自动清洗时间列（默认开启） |

### 2.5 report 专属参数
> `--time_col` / `--value_col` / `--x_col` / `--y_col` / `--sort_col` **均为可选**；
> 不传时工具会自动探测（时间列=可解析为日期的列，数值列=列名含销售/金额/额等，分类列=取值数 2~20 且列名含区域/类别等），
> 并在返回的 `result.auto_detected` 中回显探测结果。

| 参数 | 类型 | 说明 |
|------|------|------|
| `--format` | str | `json`（默认）/ `md`；`md` 时 stdout 直接输出可交付的 Markdown 正文 |
| `--chart_type` | str | 报告内图表类型，默认 bar |
| `--sort_col` / `--sort_asc` | str / flag | 报告内排序段；report 下 `--sort_asc` 缺省为升序 |
| `--filter_condition` | str | 报告内过滤条件，例：销售额>500 |
| `--time_col` / `--value_col` | str | 报告内趋势段列名 |
| `--x_col` / `--y_col` / `--title` | str | 报告内图表列名与标题 |

### 2.6 chart 专属参数
| 参数 | 类型 | 说明 |
|------|------|------|
| `--chart_type` | str | bar / line / box / hist，默认 bar |
| `--x_col` | str | x 轴列名（bar/line 必填） |
| `--y_col` | str | y 轴数值列名（bar/line/hist 必填，box 可选） |
| `--title` | str | 图表标题 |

### 2.7 白名单管理子命令
> 这两个参数不参与分析流程，**单独使用即可**（不需要 `--file` / `--intent`）。

| 参数 | 类型 | 说明 |
|------|------|------|
| `--list-whitelist` | flag | 列出当前全部信任白名单目录 |
| `--remove-whitelist-idx` | int | 删除指定序号的白名单目录（序号取自 `--list-whitelist`） |

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

`intent=report` 时 `result` 额外包含（其余意图无这些字段）：

| 字段 | 类型 | 说明 |
|------|------|------|
| `markdown` | str | **完整报告正文**，可直接作为回复内容 |
| `summary` | str | 一句话总体结论 |
| `attachments` | list[str] | 图表绝对路径；非空时必须与 `markdown` **同一条消息**发出 |
| `delivery_checklist` | list[str] | 发送前逐条自查清单 |
| `sections` | dict | 六个分段的原始结构化结果 |
| `auto_detected` | dict | 自动探测到的 time_col / value_col / x_col / sort_col |

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
| `CHART_GEN_FAILED` | 图表生成失败 |
| `NO_NUMERIC_COL` | 无数值列可分析 |
| `COLUMN_NOT_FOUND` | 指定列不存在 |
| `PARSE_EXCEL_FAILED` | Excel 解析失败 |

## 4 调用示例
前提：运行环境已配置好 PYTHONPATH，指向项目根目录

**示例 0（首选）：一次性完整报告**

```bash
# 多意图请求一律用这一条；stdout 就是最终报告正文
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent report --format md

# 需要结构化结果（报告+附件路径+自查表）时用 json（默认）
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent report --output_json result.json
```

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

示例 4：图表展示

```bash
python src/excel_analyzer/cli.py --file tests/test_data/sample_test.xlsx --intent chart --chart_type bar --x_col "区域" --y_col "销售额" --title "各区域销售额均值"
```

## 5 Agent 调用规范（调用方必须遵守）
使用子进程 subprocess 调用 cli 脚本；必须设置环境变量PYTHONPATH为项目根目录。

仅解析 stdout 作为业务返回结果；stderr 只用于调试，不得送入大模型上下文。

**输出编码契约（已由工具内部固定，调用方无需任何编码兜底）**：

- `cli.py` 启动时会把 stdout/stderr 强制 `reconfigure(encoding="utf-8")`，因此无论宿主机代码页是
  GBK(cp936) 还是 UTF-8、无论 `PYTHONIOENCODING` 怎么设，**stdout 字节流恒为 UTF-8**。
- Python subprocess 调用必须显式按 UTF-8 解码，不要依赖平台 locale：
  `subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)`，
  并在 `env` 中带上 `PYTHONIOENCODING="utf-8"`（参考 `docs/skill.py`）。
- 重定向到文件同样按 UTF-8 读取；**禁止**用 `gbk` / `mbcs` / `locale.getpreferredencoding()` 解 stdout。
- 反面故障：按 GBK 写、按 UTF-8 解 → 中文报告整段乱码，迫使智能体临场想“先把输出写文件再读”
  这类额外步骤；属于调用契约缺失，已在本版本修复。回归测试：`python -m unittest tests.test_encoding -v`。

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

`chart` 输出为 PNG 文件，只返回文件路径和大小，不返回 base64 图像数据。

`report` 把六段拼装下沉到工具内部，并且 `chart` 段额外给出**分组数值表**（均值/合计/记录数），
使报告不依赖看图即可得出区域对比结论，同时避免智能体多轮调用与大段重复上下文。

## 8 独立测试方式
在项目根目录执行单元测试，验证工具本身功能正确性：

```bash
python -m unittest tests.test_basic -v

python -m unittest tests.test_advanced -v

python -m unittest tests.test_report -v

python -m unittest tests.test_encoding -v

python -m unittest tests.test_doc_consistency -v
```

其中 `tests.test_report`（12 个用例）专门锁定「一次性完整交付」契约：
六段齐全、anomaly 四要素、chart 带 output_path 与分组数值、attachments 非空、
delivery_checklist 含「同一条消息」要求、`--format md` 输出纯 Markdown。

`tests.test_encoding`（5 个用例）锁定「stdout 恒为 UTF-8」的编码契约（见第 5 章）。

`tests.test_doc_consistency`（6 个用例）锁定**文档间内容对应性**：
两份 SKILL.md、两份 skill.py 必须逐字节一致；文档声明的用例数/测试清单必须等于实际值；
README 目录结构与实际文件树一致；文档列出的 CLI 参数必须与 `cli.py` 的 argparse 定义互相覆盖。
**改文档或改代码后必须重跑本测试，否则文档会静默过期。**

## 9 结果展示规范（最高优先级）

### 9.1 首选：`intent=report` 一次性交付

> ⚠️ **多意图请求一律走 `report`，不要逐意图多次调用，也不要自己拼装报告。**

```bash
python src/excel_analyzer/cli.py --file <xlsx> --intent report --format md
```

- `--format md`：**stdout 就是最终报告正文**，直接整段作为回复；日志仍在 stderr；stdout 已固定为 UTF-8。
- ⚠️ `--format md` 时 stdout **不是 JSON**：它是 Markdown 文本，**禁止对 stdout 调用 `json.loads`**（会直接解析失败）。
- `--format md` 可与 `--output_json <file>` **并用且互不干扰**：此时 stdout 仍只输出 Markdown 正文，
  结构化结果（`result.markdown` / `result.attachments` / `result.delivery_checklist`）另写入该文件，
  调用方按 UTF-8 读取该文件即可，**不要试图从 md 模式的 stdout 里解析 JSON**。
- 例外（失败路径）：若在生成报告前就失败（文件不存在、后缀不支持、需要用户确认白名单等），
  `--format md` 的 stdout 返回的是 JSON 错误包而非 Markdown，退出码非 0；
  调用方应先看输出是否以 `{` 开头，再决定按 JSON 还是 Markdown 处理。
- `--format json`（默认）：`result.markdown` 为报告正文，`result.attachments` 为图表路径，
  `result.delivery_checklist` 为发送前自查表。
- 走 `report` 时，**不需要**再单独调用 chart 意图，也**不需要**再单独调 message 发图。

### 9.2 交付硬性契约（四道红线）

1. **正文与附件必须同在一条消息里**。把 `result.markdown` 全文写进 `message` 的 `content`，
   把 `result.attachments` 里的图片放进同一次调用的 `media`。
   ❌ 禁止把附件单独发成一条只含图片的气泡；❌ 禁止把正文与附件拆成两条消息。
2. **不得只发图表**。图只是附件，正文必须把六个分段的关键数值写全。
3. **不得输出过程旁白**。禁止出现「正在分析…」「概览正常，现在…」「异常检测完成。现在生成柱状图：」这类句子；
   等报告生成后一次性作答。
4. **发出前对照 `delivery_checklist` 逐条自查**（六段是否齐全 / 同一条消息 / anomaly 四要素 / chart 数值）。

### 9.3 分段内容要求（`report` 已自动满足）

1. **每个意图单独列出**，用二级标题区分：`## overview` / `## stats` / `## sort_filter` / `## trend` / `## anomaly` / `## chart`
2. **必须展示关键数值**，不能只说“完成”
3. **异常检测（anomaly）必须展示**：
   - 重复行数
   - 高缺失列名及缺失率
   - 每列的离群点数量（IQR + Z-Score）
4. **图表（chart）必须返回 output_path**，并给出分组数值（均值/合计/记录数）

❌ 反例：
“异常检测完成。现在生成柱状图：[图片]”

✅ 正例（`report` 的 stdout 即为该形态）：
“## anomaly · 异常检测
- 重复行数：0 行（总行数 31）
- 高缺失列：备注（48.39%）
- 销售额：IQR 离群数 2 个（7.14%），Z-Score 离群数 2 个
- 销量 / 单价：IQR 与 Z-Score 离群数均为 0 个

## chart · 图表
- output_path：`outputs/bar_销售额_20260917_145530.png`
- 分组均值：华东 19,831.20 / 华北 18,943.33 / 华中 2,895.50 …
（图片与上述全文写在同一条消息中）”

### 9.4 单意图请求

只涉及 1 个意图（例如只要 `anomaly`）时，可单独调用该意图；
**但若要附图，必须把该段文字结论与图片放入同一条消息**。

测试用例共 **46 个**（test_basic 11 + test_advanced 12 + test_report 12 + test_encoding 5 + test_doc_consistency 6），
覆盖异常输入、边界脏数据、正常业务通路、图表可选功能、数据清洗、算法增强、一次性完整交付契约、输出编码契约、文档一致性契约。
