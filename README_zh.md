# 数据迁移工具 (Data Transferring Tool)

一个强大且稳健的工具，旨在根据预定义的规则，将数据从多个源文件（`csv`, `xls`, `xlsx`）迁移至目标文件中。在迁移过程中，该工具能够优雅地处理数据冲突，并生成详细的执行报告。

## 核心功能

- **多数据源支持：** 将一个或多个源文件的数据整合并迁移至单个目标文件中。
- **智能列映射：** 通过简洁的 YAML 配置文件，使用形如 `A: C` 的直观方式将源文件的列映射到目标文件。
- **参照关系匹配：** 根据设置的“参照列”自动在目标文件中匹配与源文件对应的行数据。
- **冲突处理机制：** 当目标单元格已经存在数据，或多个源文件尝试写入同一个单元格时，可自定义处理策略（如 `keep_original` 保留原数据, `overwrite` 覆盖, `manual` 命令行手动确认）。
- **生成详细报告：** 每次执行后都会自动生成一份详尽的 `transfer_report.xlsx` 报告，清晰地列出已迁移、被跳过的数据以及冲突处理详情。
- **双模界面 (GUI / CLI)：** 既支持通过命令行终端快速执行，也支持启动简洁的图形界面（GUI）通过鼠标点选配置文件。
- **非破坏性操作：** 严格保证绝不修改原始的源文件和目标文件，所有迁移结果均输出为全新文件。

## 安装与初始化

本项目使用 `uv` 作为依赖管理工具。

1. 确保您的系统已安装 `uv`。
2. 在终端中进入本项目的根目录。
3. 安装并同步依赖项：
   ```bash
   uv sync
   ```

## YAML 配置指南

数据迁移的规则全部在 YAML 文件中定义。以下是配置示例：

```yaml
# 目标文件路径 (支持 csv, xls, xlsx)
target_file: "target.xlsx"

# 若目标文件是 Excel，可指定要写入的工作表名或从 0 开始的 sheet 索引
target_sheet: "汇总"

# 最终生成的新文件路径（不会修改原文件）
output_file: "output.xlsx"

# 是否生成 transfer_report.xlsx
# 默认关闭；如需生成请显式设置为 true
generate_transfer_report: false

# 是否生成 reference_report.md
# 默认关闭；如需生成请显式设置为 true
generate_reference_report: false

# 是否根据冲突次数给输出 Excel 中发生冲突的单元格着色
# 默认关闭；仅对 xls / xlsx 输出生效
highlight_conflict_cells: false

# 冲突处理策略（当目标单元格已有数据时如何处理）
# 选项: 
#   - keep_original: 优先保留目标文件中已有的（或最先被写入的）数据
#   - overwrite: 后续写入的数据直接覆盖之前的数据
#   - manual: 命令行中进行交互式手动确认
conflict_resolution: "keep_original"

# 来源文件列表
sources:
  - file_path: "source1.csv"
    # 行与行之间寻找对应关系的参照列
    # 例如：将源文件的 A 列与目标文件的 C 列作为匹配条件
    reference_column:
      A: C
    
    # 需要迁移的数据列对应关系
    # 将源文件 B 列的数据写入目标文件 D 列
    # 将源文件 E 列的数据写入目标文件 F 列
    mapping:
      B: D
      E: F

  - file_path: "source2.xlsx"
    # 若来源文件是 Excel，可指定要读取的工作表名或从 0 开始的 sheet 索引
    sheet_name: "原始数据"
    reference_column:
      A: B
    mapping:
      C: D
```

- 未配置 `target_sheet` 或 `sheet_name` 时，Excel 默认使用第一个工作表。
- 当目标文件是 Excel 时，输出文件会保留目标工作簿中的其他工作表，只更新指定的目标工作表。
- 如果来源文件与目标文件是同一个 Excel，也可以通过不同的 `sheet_name` / `target_sheet` 在不同 tab 之间迁移数据。
- `generate_transfer_report` 默认是 `false`，只有显式设置为 `true` 时才会生成 `transfer_report.xlsx`。
- `generate_reference_report` 默认是 `false`，只有显式设置为 `true` 时才会生成 `reference_report.md`。
- `highlight_conflict_cells` 默认是 `false`，只有显式设置为 `true` 时才会给输出的 `xls` / `xlsx` 文件中的冲突单元格着色；`csv` 不支持此功能。
- 如果一个来源列需要写入多个目标列，推荐使用 YAML 列表写法来显式保留顺序，例如：

```yaml
mapping:
  - E: AZ
  - E: BA
  - F: BB
```

- 以上写法表示把来源 `E` 列同时写入目标 `AZ` 和 `BA` 两列，再把来源 `F` 列写入目标 `BB` 列。
- 当前版本也兼容重复键写法，但列表写法更清晰，也更符合 YAML 对“一对多映射”的表达方式。

## 使用方法

### 命令行模式 (CLI)

通过传入你的 YAML 配置文件路径来直接运行数据迁移：

```bash
uv run python main.py run path/to/config.yaml
```

### 图形界面模式 (GUI)

如果不习惯使用命令行，可以启动图形界面，通过文件浏览器选择您的配置文件并执行：

```bash
uv run python main.py gui
```

### 第三方开源许可声明 (Third-Party Notices)

您可以通过运行以下命令在终端打印出本项目使用的所有第三方开源库及其许可声明：

```bash
uv run python main.py --third-party-notices
```

## 执行报告

执行完成后，如果 YAML 中将 `generate_transfer_report` 设置为 `true`，工具会在当前目录下生成一份名为 `transfer_report.xlsx` 的 Excel 报告。报告包含了以下重要信息：
- 冲突处理结果（例如：`transferred` 已转移, `identical_skipped` 数值或文本等价因此跳过, `conflict_kept_original` 冲突-保留原数据, `conflict_overwritten` 冲突-覆盖, `skipped_not_in_target` 找不到对应行被跳过）。
- 涉及的源文件和目标文件路径。
- 涉及的源工作表和目标工作表。
- 被匹配到的参照值。
- 操作的具体目标列。
- 原数据和新数据的对比。
- 原数据与新数据的**文本相似度分数 (Similarity Score)**（当发生冲突时生成，基于模糊匹配算法，帮助您判断是否属于笔误或微小差异）。
- 当原数据与新数据只是数值表示形式不同但语义相等时（例如 `0` 和 `0.0`），工具不会覆盖目标值，而是记录为 `identical_skipped`。

如果 YAML 中将 `generate_reference_report` 设置为 `true`，工具还会在当前目录下生成一份名为 `reference_report.md` 的文本报告。该报告会按来源文件分组，列出所有**没有被 `skipped_not_in_target` 跳过**的 `reference_value`。

如果 YAML 中将 `highlight_conflict_cells` 设置为 `true`，并且输出格式为 `xls` 或 `xlsx`，工具会根据同一个目标单元格发生冲突的次数着色：
- 第 1 次冲突：黄色
- 第 2 次冲突：橙色
- 第 3 次及以上：红色
- 当原数据与新数据只是数值表示形式不同但语义相等时（例如 `0` 和 `0.0`），工具不会覆盖目标值，而是记录为 `identical_skipped`。
