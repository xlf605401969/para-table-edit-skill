---
name: parameter-table-editor
description: Edit and validate ParameterTable.xlsx-style parameter workbooks with openpyxl. Use when Codex needs to add or remove Settings-sheet macro definitions, add or remove parameter pages, add, edit, or remove parameter rows, or validate numbering and page consistency rules in a ParameterTable.xlsx file.
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
- `list-pages [--json]`
- `list-parameters --sheet PAGE [--columns id,alias,default|A,O,N ...] [--json]`
- `get-parameter --sheet PAGE (--id XXX-YYY | --alias ALIAS) [--json]`
- `add-parameter --sheet PAGE [--prefix XXX] [--suffix auto|YYY] [--json JSON_OR_@FILE] [--set key=value ...]`
- `edit-parameter --sheet PAGE (--id XXX-YYY | --alias ALIAS) [--json JSON_OR_@FILE] [--set key=value ...]`
- `remove-parameter --sheet PAGE (--id XXX-YYY | --alias ALIAS)`
- `validate`

## Editing Rules

- Treat `Settings` as special: only macro operations apply there, and macro data lives in columns `F:H`.
- Treat every non-`Settings` sheet as a parameter page.
- Use `list-pages` to query all parameter page names.
- Use `list-parameters --sheet PAGE` to query the parameters in one page; default output includes each row's `id`, `alias`, and `default`.
- Use `list-parameters --columns ...` to choose the displayed fields. Columns may be Excel letters (`A,O,N`) or friendly names (`id,alias,default`).
- Use `get-parameter --sheet PAGE --id ...` or `--alias ...` to query one parameter and print all available row properties.
- Query commands accept `--json` to return machine-readable JSON instead of text.
- Preserve the page header rows `1:3` and the `Page End` row in column `A`.
- When creating an empty page, copy the template sheet structure and remove only parameter rows.
- When adding a parameter, use `--prefix` if the page currently has no parameters. Otherwise infer the page prefix from existing rows.
- Pass row data with `--json` or repeated `--set` flags. Keys may be Excel columns (`A`, `O`, `Q`) or friendly names such as `alias`, `default`, `unit`, `desc`.
- When editing or removing a parameter, identify the target with exactly one of `--id` or `--alias`.
- Editing preserves the current row and only updates the fields provided through `--json` or `--set`.
- Editing parameter id / column `A` is intentionally rejected; use the existing id or alias only as a selector.
- New parameters always start from this default attribute set unless explicitly overridden: `hidden=0`, `read_only=0`, `storage=1`, `reset=0`, `runtime_write=0`, `limit=1`, `decimals=0`, `signed=0`, `float=0`, `reserved=0`, `max=65535`, `min=0`, `default=100`.
- New parameters must provide `alias`/`O`.

## Column Map

Main parameter columns:

- `A`: `id` / `parameter_id`
- `B`: `hidden`
- `C`: `read_only` / `readonly`
- `D`: `storage` / `store`
- `E`: `reset`
- `F`: `runtime_write`
- `G`: `limit`
- `H`: `decimals` / `decimal`
- `I`: `signed`
- `J`: `float`
- `K`: `reserved`
- `L`: `max`
- `M`: `min`
- `N`: `default`
- `O`: `alias`
- `P`: `name`
- `Q`: `unit`
- `R`: `desc` / `description`
- `S`: `calculated_default` / `computed_default`

Usage notes:

- `add-parameter` and `edit-parameter` accept either Excel columns or friendly names in `--set` and edit JSON payloads.
- `list-parameters --columns ...` accepts either Excel columns or friendly names.
- `get-parameter` returns all existing columns on the row, including columns after `S` when present.
- Default list output is `A/O/N`, which maps to `id/alias/default`.

## Typical Calls

```bash
python scripts/parameter_table_editor.py ParameterTable.xlsx validate
python scripts/parameter_table_editor.py ParameterTable.xlsx list-pages
python scripts/parameter_table_editor.py ParameterTable.xlsx list-pages --json
python scripts/parameter_table_editor.py ParameterTable.xlsx list-parameters --sheet MOTOR0
python scripts/parameter_table_editor.py ParameterTable.xlsx list-parameters --sheet MOTOR0 --columns id,alias,unit,default
python scripts/parameter_table_editor.py ParameterTable.xlsx list-parameters --sheet MOTOR0 --columns id,alias,unit,default --json
python scripts/parameter_table_editor.py ParameterTable.xlsx get-parameter --sheet MOTOR0 --id 001-000
python scripts/parameter_table_editor.py ParameterTable.xlsx get-parameter --sheet MOTOR0 --id 001-000 --json
python scripts/parameter_table_editor.py ParameterTable.xlsx add-macro --name USER_FOO --value 1
python scripts/parameter_table_editor.py ParameterTable.xlsx add-page --page-name TEST --page-prefix 13 --template-sheet BASE
python scripts/parameter_table_editor.py ParameterTable.xlsx add-parameter --sheet TEST --prefix 13 --set alias=Foo --set name=Foo --set default=0 --set max=100 --set min=0
python scripts/parameter_table_editor.py ParameterTable.xlsx edit-parameter --sheet TEST --id 013-000 --set alias=Foo2 --set default=5
python scripts/parameter_table_editor.py ParameterTable.xlsx edit-parameter --sheet TEST --alias Foo2 --set desc="updated description"
```

## Bundled Asset

Use `assets/examples/ParameterTable.xlsx` as the canonical example workbook when you need a realistic sample for validation, demos, or forward-testing.
