# Utility Bill Calculator · 水电燃气计算器

一个本地运行的 Python 桌面程序，按电、水、燃气三个分类读取 CSV / Excel 用量明细，计算每天及所选期间的费用。费率可手动修改或通过 JSON 文件导入；不需要租房合同、个人资料或外部账号。

## 功能

- 分别选择三个数据文件夹，递归扫描；也可直接添加文件。
- CSV 多选，Excel 按工作表多选，支持跨分类一起计算。
- 手动编辑用量单价、每日固定费、数据单位、计价单位、采样间隔和币种。
- JSON 费率导入；可自定义列名映射并查看工作表第一行列名。
- 起止日期筛选、每日明细、分类汇总和 CSV 导出。
- 相同记录去重、冲突检测、可选的全天数据覆盖检查。
- 保存本地设置。**不包含电费分摊功能。**
- 中英文即时切换：默认中文，点击右上角 **English** 切换英文，点击 **中文** 切回。

## 界面语言 / Interface language

每次启动默认显示中文。点击右上角 **English / 中文** 按钮可切换界面、分类名称、提示信息、数据检查和导出 CSV 的表头与说明，切换不会清空文件选择、费率、列名映射或已算出的结果。用户输入的列名、文件名、表号、单位标识和费率 JSON 字段名不翻译。操作系统文件选择窗口中的原生按钮仍由系统语言决定。

The app starts in Chinese. Click **English** in the top-right corner to switch to English, or **中文** to switch back. Your selected files, rates, column mappings and calculated amounts remain unchanged. Exported CSV headers and notes follow the current interface language. Imported names and data are never translated. Native file-dialog controls follow your operating system language.

## 安装与启动

需要 Python **3.10 或更新版本**及 Tkinter。Windows 官方 Python 安装包通常自带 Tkinter。Linux 若缺少 Tkinter，需通过系统包管理器安装（如 Ubuntu 的 `python3-tk`）。当前在 Windows / Python 3.11 验证；其他桌面系统尚未实机验证。

进入本项目文件夹后运行：

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe app.py
```

Windows 完成安装后也可双击 `start.bat`，优先使用本项目虚拟环境。

macOS / Linux：

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
```

CSV 读取和计算本身只使用 Python 标准库。Excel 读取依赖 [python-calamine](https://github.com/dimastbk/python-calamine)，无需安装 Microsoft Excel。程序不会运行 Excel 宏，也不会上传数据。

## 使用步骤

1. 将用量文件分别放入 `data/electricity/`、`data/water/`、`data/gas/`，或通过界面选择自己的文件夹。
2. 点击“刷新文件夹”。切换电费、水费、燃气费标签，按住 Ctrl / Shift 多选文件或工作表。三个标签的选择会同时保留。刷新会清除选择及临时添加的文件列表。
3. 填写所选分类的用量单价和每日固定费，或点击“导入费率 JSON”。新安装时价格为空，必须填写；没有某项费用时输入 `0`。
4. 设置文件里的数据单位，以及单价对应的计价单位。例如 CSV 以 kL 记录、水价按每 L 收费，应选择数据单位 `kL`、计价单位 `L`。
5. 如列名不同，填写“列名映射”。可先选中一张表，点击“查看选中表的列名”。
6. 可填写开始、结束日期，格式 `YYYY-MM-DD`，包含两端；留空则使用所选数据的日期范围。
7. 点击“计算选中数据”，查看每日明细、数据警告、分类总额和总计。
8. 点击“导出结果 CSV”保存结果；点击“保存设置”记住费率、币种、文件夹和列名映射。

输入、映射、日期或选择发生变化后，旧计算结果会清除，需要重新计算。保存的设置不保存文件选择及日期筛选。首次试用可以直接添加 `examples/usage.example.csv` 并导入 `examples/rates.example.json`；示例都是虚构数据，**示例费率不是任何供应商报价**。该示例电费总额为 **2.21 AUD**，数据覆盖警告是预期行为。

## 用量数据格式：CSV / Excel

支持 `.csv`、`.xlsx`、`.xls`、`.xlsm`、`.xlsb`。扩展名大小写均可；Excel 文件中的各个工作表会分别列出。**支持这些文件格式不等于识别任意排版账单**：每个被选择的工作表都必须是一张有列名的明细表。

- **第一行必须是列名**；列顺序任意，可包含额外列。不要添加顶部说明、合并表头、页脚或合计行。
- CSV 必须使用逗号分隔，编码为 UTF-8（推荐，可含 BOM）或 GB18030；单元格内逗号需使用标准 CSV 双引号转义。
- 空白数据行会跳过；空文件、缺列、无效日期或非数字用量会显示带文件名和行号的错误。
- 文件名没有限制，不必采用特定日期命名；费用日期以实际数据列为准。
- 加密 Excel、PDF、图片、HTML 伪装的 `.xls` 不支持。建议将公式转为值；程序不重新计算公式，依赖工作簿保存的计算结果，无缓存值的必填单元格会报错。

### 默认列名

| 列名 | 必填 | 含义与示例 |
| --- | --- | --- |
| `ReadDatetime` | 是 | 当地日期时间，例如 `2026-01-01T00:05:00`、`2026-01-01 00:05:00`、`2026-01-01`；也接受 Excel 日期类型单元格 |
| `ReadValue` | 是 | **该时段用量**，例如 `0.125`；必须是有限、非负数字，不带单位或千位逗号 |
| `MeterSerial` | 是 | 表号，例如 `DEMO-METER`；建议设为文本，保留前导零 |
| `MeterCode` | 否 | 通道标识，例如 `E1`；缺省视为空通道 |
| `QualityMethod` | 否 | 质量标识；`A` 视为正常，其他值或缺失会提示，但仍参与计算 |

```csv
MeterSerial,ReadDatetime,ReadValue,MeterCode,QualityMethod
DEMO-METER,2026-01-01T00:00:00,0.10,E1,A
DEMO-METER,2026-01-01T00:05:00,0.20,E1,A
```

`ReadValue` **不能是累计电表/水表读数，也不能是瞬时功率**。程序不会自动相减、也不会把 kW 当成 kWh。仅按天收取设备费时，仍需要上述三列；可为用量填写 `0`，程序根据记录日期去重计天。

日期没有时区转换；带时区的文本会拒绝，请先转换为账单当地时间。Excel 无日期格式的纯数字序列号不支持。时间戳直接归到其标记日期：如果供应商用次日 00:00 标记前一天最后一个区间的结束时间，应在导入前调整时间戳。

### 自定义列名映射

例如源表的列名是 `时间,用量,表号,通道,质量`，在对应分类中设置：

| 界面字段 | 填入的实际列名 |
| --- | --- |
| 日期时间 | `时间` |
| 时段用量 | `用量` |
| 表号 | `表号` |
| 通道 | `通道`（没有可留空） |
| 质量 | `质量`（没有可留空） |

匹配区分大小写，列名前后空格会忽略。必填字段不能为空，不能将两个字段映射到同一列。**同一分类本次所选的所有文件 / 工作表使用同一套映射和单位**；格式不一致时请分批计算。参见 `examples/custom-columns.example.csv`。

## 费率导入格式

参考并修改 `examples/rates.example.json`，然后点击“导入费率 JSON”。一次导入完整的三个分类；整个文件验证通过后才更新界面，错误不会部分应用。

顶层格式：

```json
{
  "version": 1,
  "currency": "AUD",
  "categories": {
    "electricity": {
      "usage_rate": "0.30",
      "daily_rate": "1.00",
      "data_unit": "kWh",
      "rate_unit": "kWh",
      "interval_minutes": 5
    },
    "water": {
      "usage_rate": "0.02",
      "daily_rate": "0.50",
      "data_unit": "L",
      "rate_unit": "L",
      "interval_minutes": 0
    },
    "gas": {
      "usage_rate": "0",
      "daily_rate": "0.60",
      "data_unit": "MJ",
      "rate_unit": "MJ",
      "interval_minutes": 0
    }
  }
}
```

| 字段 | 说明 |
| --- | --- |
| `version` | 必须为整数 `1` |
| `currency` | 三字母币种标签，例如 `AUD`、`CNY`；不进行汇率转换，三个分类必须使用同一币种 |
| `categories` | 必须含且仅含 `electricity`、`water`、`gas` |
| `usage_rate` | 每个计价单位的价格，非负；推荐用字符串保存小数精度 |
| `daily_rate` | 每日固定费，非负 |
| `data_unit` | 用量文件的实际单位 |
| `rate_unit` | 用量单价的计价单位 |
| `interval_minutes` | 可省略，默认 `0`；`0` 不检查全天覆盖，正整数需能整除1440，例如5、15、30、60、1440 |

单位支持 `kWh`、`MJ`、`L`、`kL`、`m³`。允许体积单位之间转换，以及 `1 kWh = 3.6 MJ` 的能量转换。**不能自动将燃气 m³ 转为 MJ**，这需要具体燃气换算系数；应先在源数据中完成转换。

热水与冷水若对同一用量叠加收费，可手动填写两项单价之和；若两者是独立计量，则分批计算。多个每日固定费可以填写合计值，但同一项费用不应重复分配到不同分类。只有设备费时，把用量单价设为 `0`。请输入含适用税费的最终价格；程序不额外加税或推算折扣。

## 计算规则与限制

```text
每日用量 = 当天去重后的区间用量之和 × 单位换算系数
每日用量费 = 每日用量 × 用量单价
每日总费 = 每日用量费 + 每日固定费
期间总费 = 所选日期范围内的每日总费之和
```

- 每个分类、每个有记录的日期计一次固定费，即使该日数据不完整。没有记录的日期不计费。因此缺少记录时，结果不是完整账单。
- 以“表号 + 通道 + 日期时间”为去重键。相同键且相同用量只算一次；相同键用量不同会报错。重复数据的质量标记会合并检查。
- 同一分类所选日期范围内只接受一个表号 / 通道，多个表请分别计算，避免固定费归属不明确。
- 当设置了采样间隔，检查普通24小时日期中从00:00起的采样点。间隔结束时标记、夏令时23/25小时日期可能出现警告；警告不改变用量加总。`0` 明确表示未验证全天覆盖。
- 中间计算使用 `Decimal`；每日展示总额与最终总额分别四舍五入至两位。最终总额基于未舍入金额求和，因此可能与逐日显示金额之和相差几分。CSV 保留每日未舍入金额。
- 当前是固定单价模型，不支持分时电价、阶梯价格、累计表数、历史费率自动切换或月度固定费用。费率变化请分批计算。
- 工作簿会读入内存，大文件计算可能暂时阻塞界面；适合日常家庭用量明细。

## 项目结构

```text
app.py                      桌面界面
billing.py                  文件导入、费率校验、计算核心
i18n.py                     中英文翻译与可翻译错误信息
requirements.txt            运行依赖
requirements-dev.txt        测试依赖
start.bat                   Windows 启动入口
examples/                   虚构用量与费率示例
data/electricity/            本地电费数据
data/water/                  本地水费数据
data/gas/                    本地燃气数据
exports/                    本地导出结果
tests/                      测试
settings.local.json          运行后保存的本地设置（不纳入版本控制）
```

## 测试

```powershell
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

测试使用虚构数据，覆盖 CSV、XLSX、XLSM、XLS、列名映射、日期筛选、单位换算、去重、冲突、费率校验和窗口交互。GUI 测试需要桌面显示环境。XLSB 使用相同 Calamine 读取后端支持，尚未添加端到端 XLSB 测试样本。

## 许可证

[MIT](LICENSE)。
