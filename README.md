# Data Transferring Tool

A powerful, robust tool to transfer data from multiple source files (`csv`, `xls`, `xlsx`) to a target file based on predefined rules, while handling conflicts gracefully and generating detailed reports.

## Features

- **Multiple Sources:** Transfer data from one or more source files to a single target file.
- **Smart Mapping:** Map columns between source files and target files intuitively using YAML config.
- **Reference Resolution:** Looks up corresponding rows in the target file based on a reference column.
- **Conflict Handling:** Handle data conflicts when target cells already have data or multiple sources write to the same cell (`keep_original`, `overwrite`, `manual`).
- **Reporting:** Generates a comprehensive `transfer_report.xlsx` detailing what was transferred, skipped, or had conflicts.
- **GUI and CLI:** Use it in terminal or launch a simple GUI to select the config file.
- **Non-Destructive:** Does not modify original source or target files.

## Installation & Setup

This project uses `uv` for dependency management.

1. Ensure `uv` is installed on your system.
2. Navigate to the project directory.
3. Install dependencies:
   ```bash
   uv sync
   ```

## Configuration YAML Format

The data transfer rules are defined in a YAML file. Here is an example:

```yaml
# Target file path (supports csv, xls, xlsx)
target_file: "target.xlsx"

# The output file where the merged data will be saved
output_file: "output.xlsx"

# How to handle conflicts when a target cell already has data
# Options: 
#   - keep_original: Keep the existing data in the target
#   - overwrite: Later data overwrites earlier data
#   - manual: Prompt for manual confirmation in CLI
conflict_resolution: "keep_original"

# List of source files
sources:
  - file_path: "source1.csv"
    # Reference columns to match rows between source and target
    # Here, Source Column A corresponds to Target Column C
    reference_column:
      A: C
    
    # Columns to transfer from source to target
    # Source Column B maps to Target Column D
    # Source Column E maps to Target Column F
    mapping:
      B: D
      E: F

  - file_path: "source2.xlsx"
    reference_column:
      A: B
    mapping:
      C: D
```

## Usage

### Command Line Interface (CLI)

Run the tool by passing your YAML configuration file:

```bash
uv run python main.py run path/to/config.yaml
```

### Graphical User Interface (GUI)

Launch the GUI to select your configuration file interactively:

```bash
uv run python main.py gui
```

### Third-Party Notices

You can print the licenses of all third-party dependencies used in this project by running:

```bash
uv run python main.py --third-party-notices
```

## Packaging as an Executable (PyInstaller)

You can package the tool into a standalone executable so that it can be run without installing Python or any dependencies.

### Command to package (with CLI and Third-Party Notices included)

Run the following command. The `--add-data` flag ensures the `ThirdPartyNotices.txt` is bundled correctly into the executable.

```bash
# On macOS / Linux
uv run pyinstaller --onefile --add-data "ThirdPartyNotices.txt:." main.py

# On Windows
uv run pyinstaller --onefile --add-data "ThirdPartyNotices.txt;." main.py
```

Once the build is complete, you will find the executable file inside the `dist/` folder. You can then run it directly and use all the options, including:

```bash
./dist/main --third-party-notices
./dist/main run config.yaml
```

## Report

After execution, the tool generates a `transfer_report.xlsx` containing:
- Conflict resolution method used (e.g., `transferred`, `conflict_kept_original`, `conflict_overwritten`, `skipped_not_in_target`).
- Source and Target file paths.
- Reference values used to match the rows.
- The columns affected.
- Original data vs New Data.
- Similarity Score between the old and new data (if a conflict occurred).
