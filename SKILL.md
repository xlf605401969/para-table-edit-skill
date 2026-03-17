---
name: parameter-table-editor
description: Edit and validate ParameterTable.xlsx-style parameter workbooks with openpyxl. Use when Codex needs to add or remove Settings-sheet macro definitions, add or remove parameter pages, add or remove parameter rows, or validate numbering and page consistency rules in a ParameterTable.xlsx file.
---

# Parameter Table Editor

Use `scripts/parameter_table_editor.py` for all workbook changes and validation. Prefer running the script over hand-editing cells so numbering, styles, and page structure stay consistent.

## Workflow

1. Inspect the workbook with `validate` before editing if the file may already be inconsistent.
2. Apply the requested mutation with one of the edit subcommands.
3. Run `validate` again after every structural change.
4. Read `references/format-rules.md` when you need the full page structure, column mapping, and validation contract.
5. Report both hard errors and continuity warnings back to the user.

## Commands

Run the script with:

```bash
python /path/to/parameter-table-editor/scripts/parameter_table_editor.py <workbook> <command> ...
```

Supported commands:

- `add-macro --name NAME --value VALUE [--description TEXT]`
- `remove-macro --name NAME`
- `add-page --page-name NAME --page-prefix XXX [--template-sheet BASE] [--copy-parameters]`
- `remove-page --page-name NAME`
- `add-parameter --sheet PAGE [--prefix XXX] [--suffix auto|YYY] [--json JSON_OR_@FILE] [--set key=value ...]`
- `remove-parameter --sheet PAGE (--id XXX-YYY | --alias ALIAS)`
- `validate`

## Editing Rules

- Treat `Settings` as special: only macro operations apply there, and macro data lives in columns `F:H`.
- Treat every non-`Settings` sheet as a parameter page.
- Preserve the page header rows `1:3` and the `Page End` row in column `A`.
- When creating an empty page, copy the template sheet structure and remove only parameter rows.
- When adding a parameter, use `--prefix` if the page currently has no parameters. Otherwise infer the page prefix from existing rows.
- Pass row data with `--json` or repeated `--set` flags. Keys may be Excel columns (`A`, `O`, `Q`) or friendly names such as `alias`, `default`, `unit`, `desc`.
- New parameters always start from this default attribute set unless explicitly overridden: `隐藏=0`, `只读=0`, `存储=1`, `复位=0`, `运行时写入=0`, `限制=1`, `小数点=0`, `符号=0`, `浮点=0`, `备用属性=0`, `最大值=65535`, `最小值=0`, `默认值=100`.
- New parameters must explicitly provide both `alias`/`O` and `name`/`P`; the script rejects rows that leave either field empty.

## Typical Calls

```bash
python scripts/parameter_table_editor.py ParameterTable.xlsx validate
python scripts/parameter_table_editor.py ParameterTable.xlsx add-macro --name USER_FOO --value 1
python scripts/parameter_table_editor.py ParameterTable.xlsx add-page --page-name TEST --page-prefix 13 --template-sheet BASE
python scripts/parameter_table_editor.py ParameterTable.xlsx add-parameter --sheet TEST --prefix 13 --set alias=Foo --set name=Foo --set default=0 --set max=100 --set min=0
```

## Bundled Asset

Use `assets/examples/ParameterTable.xlsx` as the canonical example workbook when you need a realistic sample for validation, demos, or forward-testing.
