# Format Rules

Use this file when you need the detailed workbook structure and validation contract for `ParameterTable.xlsx`.

## Workbook Layout

- `Settings` is the only non-parameter sheet.
- Macro definitions live in `Settings!F:H`.
- Every other sheet is a parameter page.
- Parameter pages keep fixed header rows `1:3`.
- Parameter data starts at row `4`.
- Each parameter page ends with a `Page End` marker in column `A`.

## Parameter Row Shape

Core columns:

- `A`: parameter id, formatted as `XXX-YYY`
- `B`: hidden
- `C`: read only
- `D`: storage
- `E`: reset
- `F`: runtime write
- `G`: limit
- `H`: decimals
- `I`: signed
- `J`: float
- `K`: reserved
- `L`: max
- `M`: min
- `N`: default
- `O`: alias
- `P`: name
- `Q`: unit
- `R`: description
- `S`: calculated default

Columns after `S` may exist on some sheets and should be preserved.

## New Parameter Defaults

When adding a new parameter, initialize these defaults unless the user explicitly overrides them:

- hidden: `0`
- read only: `0`
- storage: `1`
- reset: `0`
- runtime write: `0`
- limit: `1`
- decimals: `0`
- signed: `0`
- float: `0`
- reserved: `0`
- max: `65535`
- min: `0`
- default: `100`

The new row must also provide:

- `O` / `alias`
- `P` / `name`

## Validation Rules

Hard errors:

- Header markers in rows `1:3` do not match the expected template.
- A parameter row is missing column `A`.
- Column `A` is not in `XXX-YYY` format.
- One sheet mixes different `XXX` prefixes.
- A `YYY` suffix repeats within the same sheet.
- A `YYY` suffix exceeds `127`.
- Two different sheets reuse the same `XXX` prefix.
- Column `O` (`Alias`) is empty.

Warnings:

- A parameter page has no parameter rows.
- Parameter suffixes are not continuous.
- Parameter suffixes do not start at `000`.
