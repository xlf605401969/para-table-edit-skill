# parameter-table-editor

Codex skill for querying, editing, and validating `ParameterTable.xlsx`, including page/parameter listing, per-parameter detail lookup, and add/edit/remove operations for parameter rows.

Repository layout:

- `SKILL.md`: skill definition
- `agents/openai.yaml`: UI metadata
- `scripts/parameter_table_editor.py`: workbook edit and validation CLI
- `references/format-rules.md`: workbook structure and validation contract
- `assets/examples/ParameterTable.xlsx`: example workbook for testing and reference

Only the bundled example workbook is tracked. Other local `.xlsx` files stay ignored by `.gitignore`.
