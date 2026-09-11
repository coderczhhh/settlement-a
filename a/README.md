# A 组：用户侧售电信息分类汇总

本目录实现 A 组清算数据处理流程：读取“用户侧清算数据”和“永强供电公司月度数据”两个 Excel 工作簿，标准化业务明细，完成市场主体与用户电量汇总、勾稽检查，并输出 JSON 结果及 Excel 人工核对表。

> 本公开仓库只包含程序代码、环境说明和测试代码，不包含任何生产或测试业务数据。请自行准备符合下述格式的输入文件。

## 运行环境

- Python 3.11
- `openpyxl >= 3.1, < 4`
- Conda（Miniconda 或 Anaconda）

### 新建 `settlement` Conda 环境

在仓库根目录执行：

```bash
conda env create -f a/environment.yml
conda activate settlement
```

如果本机已经存在 `settlement` 环境，可更新依赖：

```bash
conda env update -n settlement -f a/environment.yml --prune
conda activate settlement
```

也可以手动创建：

```bash
conda create -n settlement python=3.11 -y
conda activate settlement
python -m pip install -r a/requirements.txt
```

验证安装：

```bash
python -c "import openpyxl; print(openpyxl.__version__)"
```

## 输入格式

程序接收两个 `.xlsx` 文件和一个可选数值参数。文件名需分别包含 `用户侧清算数据` 和 `永强供电公司`，程序据此识别工作簿。

### 1. 用户侧清算数据 Excel

必须包含以下工作表：

| 工作表 | 表头要求 | 用途 |
| --- | --- | --- |
| `1.结算_用户属性用户侧售电信息` | 前 20 行内须出现：`用电分类`、`分类明细`、`电压等级`、`市场主体类型`、`分时类型`、`售电电量`、`电能电费` | 用户侧售电量、电费及市场主体汇总 |
| `4.结算_1.5倍代理购电信息` | 前 20 行内须有一行满足 A 列为 `户号`、D 列为 `电压等级`；该表头占 3 行，F 列为电量合计，N 列为电能电费合计 | 追加到代理购电用户 |

### 2. 永强供电公司月度数据 Excel

必须包含 `luxi` 工作表，且前 20 行内须出现以下表头：

```text
用电分类、 电压等级、 市场主体类型、 分时类型、 售电电量
```

`Sheet3` 可存在，但仅由 A1 原样读取，不参与 A2-A5 汇总，避免与 `luxi` 重复计算。

### 3. 抽水蓄能参数

`--pumped-storage-kwh` 表示抽水蓄能用电量，单位为千瓦时。该数据不从上述两个 Excel 自动提取，默认值为 `0`。

### 输入约束

- Excel 公式单元格使用 `data_only=True` 读取已缓存数值；程序不计算公式。输入文件应先由 Excel/WPS 完成公式计算并保存。
- 空数值按 `0` 处理；非法数值、布尔值、无穷大或 NaN 会报错。
- 输入文件不能是名称以 `~$` 开头的 Excel 临时锁定文件。
- 未识别的市场主体类型或电压等级会直接终止流程，防止静默漏算。
- 电量单位为千瓦时，电费单位为元。

## A 流程执行逻辑

```text
两个 Excel 文件
    │
    ▼
A1 读取全部 sheet 的已计算单元格值
    │  a_1_all_workbooks.json
    ▼
A2 识别表头、清洗字段、提取三类标准明细及排除记录
    │  a_2_standardized_records.json
    ├───────────────────────┐
    ▼                       ▼
A3 市场主体×电压等级汇总   A4 用户电量汇总（不含/含永强）
    │                       │
    └───────────┬───────────┘
                ▼
A5 合成最终 JSON 并检查勾稽关系
                │  a_5_final_table.json
                ▼
A6 生成静态数值 Excel 核对表
                │  a_6_statistics_comparison.xlsx
                ▼
A0 写入运行清单 a_0_run_manifest.json
```

各步骤职责如下：

| 步骤 | 模块 | 处理逻辑 |
| --- | --- | --- |
| A1 | `a_1_read_excel_to_dict.py` | 每个工作簿只打开一次，读取全部 sheet 和全量单元格值，保留来源结构 |
| A2 | `a_2_standardize_records.py` | 提取 `user_sales`、`proxy_15x`、`yongqiang` 三类记录，记录来源文件、sheet、行号和排除项 |
| A3 | `a_3_aggregate_market_subjects.py` | 按市场主体类型、电压等级汇总售电电量与电能电费，并与来源总量勾稽 |
| A4 | `a_4_build_user_totals.py` | 生成“不含永强”和“含永强”用户电量表，追加 1.5 倍代理购电和显式抽蓄参数 |
| A5 | `a_5_build_final_table.py` | 合并 A3、A4 结果，形成最终 JSON，总检查不通过时终止 |
| A6 | `a_6_export_json_to_excel.py` | 将最终 JSON 写成四个 sheet 的格式化 Excel，不写 Excel 公式 |
| A0 | `a_0_run_all.py` | 顺序调度 A1-A6，输出运行时间、记录数、关键结果和文件清单 |

## 业务口径

1. 用户侧表中市场主体类型 `-` 计入居民、农业。
2. `1批发市场用户` 和 `3零售用户` 合并为市场化用户。
3. `2兜底用户` 单列。
4. `4代理用户` 加上 `4.结算_1.5倍代理购电信息` 的电量，形成代理购电用户。
5. 居民、农业按 `1kV及以下 / 1-10（20）kV / 35kV及以上` 分类。
6. 工商业按 `不满1千伏 / 1-10千伏及以上` 分类。
7. 永强 `luxi` 中市场主体为空且用电分类以 `一、` 或 `四、` 开头的行计入居民农业；`代理用户`、`零售用户` 分别计入代理、市场化。
8. 永强 `Sheet3` 为辅助计算页，不重复加入汇总。
9. 抽水蓄能电量由 `--pumped-storage-kwh` 显式传入，默认 `0`。
10. 永强声明合计与有效明细不一致时标记 `WARNING`，不终止流程；核心勾稽不一致时终止。
11. `五、打水用电` 不参与 A 组汇总；A2 会把对应行写入排除清单。

## 执行代码样例

以下命令均在仓库根目录执行。

### 一键运行 A1-A6

```bash
python -m a.a_0_run_all \
  --user-file "/path/to/2026年6月用户侧清算数据.xlsx" \
  --yongqiang-file "/path/to/永强供电公司2026年06月月度数据表.xlsx" \
  --output-dir "outputs/a_group" \
  --pumped-storage-kwh 0
```

带抽水蓄能电量的示例：

```bash
python -m a.a_0_run_all \
  --user-file "/path/to/用户侧清算数据.xlsx" \
  --yongqiang-file "/path/to/永强供电公司月度数据.xlsx" \
  --output-dir "outputs/a_group" \
  --pumped-storage-kwh 125000.5
```

### Python API 调用

```python
from decimal import Decimal

from a.a_0_run_all import a_0_run_all

result = a_0_run_all(
    user_file="/path/to/用户侧清算数据.xlsx",
    yongqiang_file="/path/to/永强供电公司月度数据.xlsx",
    output_dir="outputs/a_group",
    pumped_storage_kwh=Decimal("125000.5"),
)

print(result["manifest"]["status"])
print(result["final_table"]["audit"]["status"])
```

### 单步运行

```bash
# A1：Excel -> 全量 JSON
python -m a.a_1_read_excel_to_dict \
  --user-file "/path/to/用户侧清算数据.xlsx" \
  --yongqiang-file "/path/to/永强供电公司月度数据.xlsx" \
  --output "outputs/a_group/a_1_all_workbooks.json"

# A2：标准化明细
python -m a.a_2_standardize_records \
  --input "outputs/a_group/a_1_all_workbooks.json" \
  --output "outputs/a_group/a_2_standardized_records.json"

# A3：市场主体与电压等级汇总
python -m a.a_3_aggregate_market_subjects \
  --input "outputs/a_group/a_2_standardized_records.json" \
  --output "outputs/a_group/a_3_market_subject_voltage_summary.json"

# A4：用户电量汇总
python -m a.a_4_build_user_totals \
  --input "outputs/a_group/a_2_standardized_records.json" \
  --pumped-storage-kwh 0 \
  --output "outputs/a_group/a_4_user_volume_summary.json"

# A5：合成最终 JSON
python -m a.a_5_build_final_table \
  --market-summary "outputs/a_group/a_3_market_subject_voltage_summary.json" \
  --user-volume-summary "outputs/a_group/a_4_user_volume_summary.json" \
  --output "outputs/a_group/a_5_final_table.json"

# A6：最终 JSON -> Excel 核对表
python -m a.a_6_export_json_to_excel \
  --final-json "outputs/a_group/a_5_final_table.json" \
  --standardized-json "outputs/a_group/a_2_standardized_records.json" \
  --output "outputs/a_group/a_6_statistics_comparison.xlsx"
```

## 输出格式

默认输出目录结构：

```text
outputs/a_group/
├── a_0_run_manifest.json
├── a_1_all_workbooks.json
├── a_2_standardized_records.json
├── a_3_market_subject_voltage_summary.json
├── a_4_user_volume_summary.json
├── a_5_final_table.json
└── a_6_statistics_comparison.xlsx
```

### `a_0_run_manifest.json`

运行清单，主要字段：

- `status`：流程状态，成功时为 `PASS`。
- `started_at`、`finished_at`、`elapsed_seconds`：运行时间信息。
- `runtime`：Python 和操作系统版本。
- `inputs`、`parameters`、`outputs`：输入路径、抽蓄参数和产物清单。
- `record_counts`：三类标准化记录数量。
- `key_results`：不含永强、含永强、含抽蓄的总量。
- `final_audit_status`：最终勾稽状态。

### `a_1_all_workbooks.json`

完整保留两个 Excel 的层级和值：

```json
{
  "schema_version": "1.1",
  "read_mode": "data_only_values_once",
  "workbooks": [
    {
      "file_name": "用户侧清算数据.xlsx",
      "sheet_names": ["..."],
      "sheets": {
        "工作表名": {
          "max_row": 100,
          "max_column": 20,
          "rows": [
            {"row_number": 1, "values": ["字段1", "字段2"]}
          ]
        }
      }
    }
  ]
}
```

### `a_2_standardized_records.json`

`records` 中包含 `user_sales`（用户侧售电）、`proxy_15x`（1.5 倍代理购电）和 `yongqiang`（永强有效明细）。每条明细保留来源：

```json
{
  "售电电量": 123.45,
  "来源": {
    "workbook": "用户侧清算数据.xlsx",
    "sheet": "1.结算_用户属性用户侧售电信息",
    "row": 12
  }
}
```

`excluded_rows` 记录未参与汇总的行及原因，`statistics` 记录数量、市场主体分布和永强声明合计。

### `a_3_market_subject_voltage_summary.json`

`rows` 是按市场主体类型、电压等级生成的行字典列表，字段为 `市场主体类型`、`电压等级`、`求和项:售电电量`、`求和项:电能电费`。`audit` 给出来源合计、汇总合计、差额和 `PASS/FAIL` 状态。

### `a_4_user_volume_summary.json`

- `without_yongqiang`：不含永强的 `用户类型 / 用电量` 行列表。
- `with_yongqiang`：含永强的 `用户类型 / 用电量` 行列表，并包含 `合计（含抽水蓄能）`。
- `assumptions`：抽蓄参数、1.5 倍代理追加及 Sheet3 排除口径。
- `audit`：工商业小计、总计、永强增量、抽蓄等勾稽结果。

### `a_5_final_table.json`

最终交付 JSON。`sections` 下包含 `市场主体类型与电压等级汇总`、`不含永强`、`含永强`。`audit.status` 为最终状态；任一核心检查未通过时不会正常生成最终结果。

### `a_6_statistics_comparison.xlsx`

由 JSON 生成的静态数值核对表，不包含 Excel 计算公式，包含四个工作表：

- `统计总表`：并排展示市场主体汇总、不含永强和含永强。
- `市场主体汇总`：市场主体类型与电压等级明细。
- `用户电量汇总`：不含永强与含永强对比及图表。
- `核对与排除`：勾稽差额、警告、记录数和排除行。

## 勾稽与异常处理

流程自动检查：

- 市场主体分压汇总与用户侧来源明细的总电量、总电费一致。
- 工商业小计等于代理、兜底、市场化之和。
- 总计等于居民农业与工商业之和。
- 含永强总计减不含永强总计等于永强有效明细合计。
- 含抽水蓄能合计等于含永强总计加显式抽蓄参数。
- 永强声明合计与有效明细不一致时输出 `WARNING`，但不阻断其他核心检查。

## 测试

```bash
python -m unittest discover -s a/tests -v
```

公开仓库不附带业务 Excel。若本地不存在测试夹具，集成测试会自动跳过；把脱敏测试文件放到仓库根目录的 `datafortest/` 后即可运行完整集成测试。
