# amount_unifier — 金额单位统一子任务

把收集表格中**格式不统一的金额列**，统一换算成**目标单位**（默认万元）的数值。是"表格清稿工具"的一个子任务，既可单独运行，也可作为模块被其他脚本调用。

- 脚本：`amount_unifier.py`（单文件，内部分层）
- 依赖：`openpyxl`（`pip install openpyxl`）
- 示例数据：`examples/sample_input.xlsx`（虚构）→ 结果自动生成为 `{输入}_clean_{mode}_{时间戳}.xlsx`

---

## 一、快速使用

### 1. 命令行
```bash
# 处理 H、I 两列，保留原列并在其后追加"结果列+说明列"（默认模式）
python amount_unifier.py -i input.xlsx -c H I

# 覆盖模式：直接改写原列，不生成说明列，结果文件可直接使用
python amount_unifier.py -i input.xlsx -c H I -m overwrite

# 用列名指定、指定工作表与数据行、关闭底色
python amount_unifier.py -i input.xlsx -c 金额（万元） -s "Sheet1 (2)" -r 3 111 --no-color
# 统一为“元”并输出 JSON 报告
python amount_unifier.py -i input.xlsx -c H I --target-unit 元 --json
```

### 2. 作为模块调用
```python
from amount_unifier import unify_columns, normalize_amount

# 处理整个文件的若干列，返回结构化报告（库内不打印业务日志）
report = unify_columns("input.xlsx", ["H", "I"], mode="overwrite")
print(report.total, report.success, len(report.manual))   # 218 212 6
for m in report.manual:                                    # 需人工记录
    print(m["coord"], m["rule"], m["note"])

# 复用单个值的换算（纯函数）
normalize_amount("2004279.6元", "s").value     # -> 200.428
normalize_amount("=11000000+2400000", "f", 13400000).value   # -> 13400000.0（公式如实保留）
```

### 3. 项目结构与本地运行

```
table-cleaner/
├── amount_unifier.py      # 核心模块（单文件，可独立运行/调用）
├── docs/roadmap.md        # 路线图
├── examples/              # 虚构示例数据与生成脚本
│   ├── sample_input.xlsx
│   └── make_sample.py
├── tests/                 # 单元测试（纯函数，不依赖真实数据）
│   └── test_amount_unifier.py
├── requirements.txt
└── LICENSE
```

安装依赖、运行测试、生成示例：
```bash
pip install -r requirements.txt
python -m unittest discover -s tests -v     # 运行测试
python examples/make_sample.py              # 重新生成示例数据
```

用示例数据快速体验：
```bash
python amount_unifier.py -i examples/sample_input.xlsx -c H I
```

---

## 效果展示（以 `examples/sample_input.xlsx` 为例）

对示例数据执行（append 模式，目标单位默认万元）：

```bash
python amount_unifier.py -i examples/sample_input.xlsx -c H I
```

处理前后对照（金额列 H，结果写入新增的 I 列）：

| 行 | 原始值 | 标准化结果 | 处理方式 | 底色 |
|---|---|---|---|---|
| H3 | 11.44 | 11.44 | 数值 | 🟩 |
| H4 | 1567.9341 | 1567.9341 | 数值 | 🟩 |
| H5 | 2004279.6元 | 200.428 | 文本元（÷10000） | 🟧 |
| H6 | 6581万元 | 6581 | 文本万元 | 🟦 |
| H7 | 777.3万 | 777.3 | 文本万 | 🟦 |
| H8 | 2亿元 | 20000 | 文本亿元（×10000） | 🟦 |
| H9 | 5千元 | 0.5 | 文本千元（÷10） | 🟦 |
| H10 | =11000000+2400000 | 保留原值 | 需人工·公式无缓存 | 🟥 |
| H11 | 1万-100万不等 | 保留原值 | 需人工·区间文本 | 🟥 |
| H12 | 累计赔偿限额200万，每次事故100万…… | 保留原值 | 需人工·长文本 | 🟥 |
| H13 | 973.7256万元 | 973.7256 | 文本万元（合并锚点） | 🟦 |
| H14/H15 | 合并区，继承 H13 | 973.7256 | 合并继承 | 🟦 |

要点：
- **有明确单位才换算**：`元→÷10000`、`亿元→×10000`、`千元→÷10`、`万/万元`按目标单位取值，换算基于 `UNIT_SCALE`。
- **无单位不猜测**：纯数值保守保留（H3/H4）；公式如实取计算值，若 Excel 未计算则无缓存→转人工（H10）。
- **拿不准交人工**：区间、长文本保留原值并标红（H11/H12），供人工填写。
- **合并同步**：原列 H13:H15 合并，结果列 I13:I15 也合并，仅锚点出值、其余继承，视觉与原列一致。

---

## 二、参数说明

### CLI
| 参数 | 说明 |
|---|---|
| `-i / --input` | 输入 xlsx 路径（必填） |
| `-c / --columns` | 待处理列，多个用空格分隔；支持列字母（`H I`）或列名（如 `金额（万元）`） |
| `-s / --sheet` | 工作表名；缺省取**第一张表** |
| `--header-row` | 表头行号；缺省自动探测（前 10 行里第一个"非空文本≥3"的行） |
| `-r / --rows` | 数据行范围 `START END`；缺省为 `表头行+1` 到 `末行` |
| `-o / --output` | 输出路径；缺省自动生成 `{原文件名}_clean_{mode}_{时间戳}.xlsx` |
| `-m / --mode` | `append`（默认，保留原列+结果列+说明列）\| `overwrite`（覆盖原列） |
| `--target-unit` | 目标单位：`元/千元/万元/亿元`，默认 `万元` |
| `--json` | 以 JSON 输出机器可读报告（而非人眼文本） |
| `--no-color` | 不用底色标注 |
| `--no-legend` | 不生成"颜色图例"工作表 |
| `--log-level` | `DEBUG/INFO/WARNING/ERROR`，默认 `INFO` |

### 库函数 `unify_columns(...)`
签名见脚本内文档字符串，常用键：`input_path, columns, sheet, header_row, row_start, row_end, mode, output_path, target_unit, text_unit_rules, colorize, legend, logger`。返回 `CleanReport`（含 `total/success/manual/by_rule/output_path/target_unit` 等；`to_dict()` 可转 JSON）。

---

## 三、实现逻辑

### 分层结构（单文件内）
```
规则层  AmountNormalizer / TextUnitRule   纯逻辑：值+类型 -> 目标单位数值，不依赖 Excel，可单测
编排层  AmountColumnCleaner               只负责 Excel 读写、列定位、合并、上色、两种落地模式
报告层  CleanReport                       统计与需人工清单
入口层  main(argv) / unify_columns(...)   CLI 与库 API
```

### 单值标准化规则（按优先级）
1. **空值** → 需人工
2. **公式**（`data_type=='f'`）：用 `data_only=True` 读**缓存计算值**，如实保留（不做单位换算）
3. **数值**（`data_type=='n'`）：**保守保留原值**（按目标单位口径），不做单位猜测（不因数值大而 ÷10000）
4. **文本**：按单位规则表匹配
   - `x亿元`/`x万元`/`x千元`/`x万` → 按目标单位取值；`x元` → 按目标单位换算
   - 区间（含"不等/-/至"）或长文本（>20字）→ 需人工，**保留原值**
5. 其他类型 → 需人工

> 文本单位由 `_SRC_UNITS`（源单位识别）+ `unit_rules_for(target_unit)` 生成规则表；已识别 元/千元/万元/亿元，新增单位只需加一条（见"五、扩展"）。

### 合并单元格处理（重点）
- **读取**：openpyxl 中合并区只有锚点有值，非锚点是 `MergedCell`（值为 None）。用 `merged_map` 记录"被覆盖格→锚点"，非锚点格**继承锚点**的标准化结果（rule 记为"合并继承"）。
- **append 模式**：在原列后插入"结果列+说明列"，并把原列的**单列纵向合并区复刻**到这两列——非锚点行不写值/不设色，只写锚点行再 merge，使辅助列与原列一样是**真正的合并**，视觉统一。生成列的字体、对齐、边框自动沿用目标列（表头加粗以示区分）。
- **overwrite 模式**：不动合并结构，只改锚点格的值，非锚点保持 `MergedCell`。
- **插入平移**：`insert_cols` 不会自动调整合并区，用 `_shift_merges_for_insert` 手动把插入点右侧的合并区右移，避免错位。

### 两种落地模式
| 模式 | 原列 | 结果列 | 说明列 | 用途 |
|---|---|---|---|---|
| `append`（默认） | 保留 | 原列后新增 | 原列后新增 | 核对检查 |
| `overwrite` | 覆盖 | 即原列 | 不生成 | 直接使用 |

---

## 四、颜色图例（底色标注处理类型）

| 底色 | 类型 | 含义 |
|---|---|---|
| 🟩 浅绿 | 数值 | 原值为数值 |
| 🟦 浅蓝 | 文本万/万元/千元/亿元 | 原文带中文数量单位，按单位换算 |
| 🟧 浅橙 | 文本元 | 原文带'元'单位，按单位换算 |
| 🟨 浅黄 | 公式 | 公式计算值，如实保留 |
| 🟥 浅红 | 需人工 | 区间/长文本等，保留原值 |

图例标注的是**源类型**（与目标单位无关），所有值统一换算到 `--target-unit`（默认万元），图例 sheet 顶部会写明目标单位。结果列与说明列使用**同一套配色**；合并单元格不单独配色，整块显示**锚点格**的类型色。结果列数值以数值型两位小数（`0.00`）显示，底层值保留原始精度。可用 `--no-color` 关闭。

---

## 五、扩展

**切换目标单位**：命令行 `--target-unit 元|千元|万元|亿元`（默认 `万元`），或库调用 `unify_columns(..., target_unit="元")`。换算基于单位体系表 `UNIT_SCALE`：`结果 = 数值 × 源单位比例 ÷ 目标单位比例`。**无单位的纯数值/公式仍保守保留，不参与换算。**

**新增“源单位”识别**（如“圆”=元 等同义词）：在 `_SRC_UNITS` 增加一条 `(正则, 规则名, UNIT_SCALE键)`，并确保 `UNIT_SCALE` 含该单位；默认规则表由 `unit_rules_for(target_unit)` 自动生成。目前已识别 元/千元/万元/亿元。

**完全自定义规则**：仍可注入 `text_unit_rules`（提供则优先于按目标单位生成的默认规则），此时 `convert` 需自行换算到目标单位：
```python
import re
from amount_unifier import unify_columns, TextUnitRule
rules = (TextUnitRule(re.compile(r"^\s*([\d,\.]+)\s*亿元\s*$"), "文本亿元", lambda n: n * 10000),)
unify_columns("表.xlsx", ["H"], text_unit_rules=rules)
```

**改配色**：修改文件顶部的 `FILL_*` 与 `_RULE_FILL`、`LEGEND_ITEMS`。

---

## 六、注意事项
- **公式**必须用 `data_only=True` 读缓存值；若 Excel 从未打开计算过，缓存可能为空 → 记为需人工。
- **万元特殊格式**（如 `0"."0,"万元"`）只是显示伪装，脚本以单元格**真实类型和值**判断，不看显示。
- 文本套数字格式（如 `@`、`#,##0.00`）对文本无害，按文本规则处理。
- 默认取第一张表作为工作表，并自动探测表头/数据行；如探测不准，用 `-s / --header-row / -r` 显式指定。
- 输出为**新文件**，不覆盖输入；若目标文件正被 Excel 打开会导致保存失败（PermissionError），请先关闭。
