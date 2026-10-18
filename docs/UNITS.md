# Unit table

Canonical names are ASCII (`degC`, `ug`, `mg/L`) so tool output does not
depend on a degree sign or a micro sign surviving the file encoding. The
source unit is preserved on each column as `source_unit`.

Conversion is `canonical = value * factor + offset`. Temperature needs the
offset. Everything else uses `offset = 0`.

`ppm` is the dilute aqueous approximation `1 ppm = 1 mg/L`. The column
carries `unit_note` when that alias is used. Mass-per-volume factors other
than `ppm` are exact SI conversions. `1 ug/mL = 1 mg/L`.

## Built-in dimensions

| Dimension | Canonical | Examples |
| --- | --- | --- |
| concentration | mg/L | mg/L, ug/mL, g/L, ng/mL, g/m3, ppm |
| temperature | degC | degC, C with a degree sign, K, degF |
| mass | g | g, mg, ug, ng, kg |
| volume | mL | mL, L, uL, dL |
| pressure | kPa | kPa, Pa, bar, atm, psi, mmHg |
| length | m | m, cm, mm, um, nm, km |
| time | s | s, ms, min, h |
| fraction | % | %, percent |

Bare `C` and bare `F` are not treated as temperature. Bare `m` is metres,
not minutes (`min`).

Headers are split when the unit is the whole cell (`mg/L` on its own header
row) or when it sits in parentheses or brackets (`Conc. (ug/mL)`,
`Temp [K]`). A parenthetical word such as `(replicate)` stays part of the
name.

## Adding a unit

In process, for tests and embedding:

```python
from mcp_advanced_excel.units import register_unit

register_unit("RFU", canonical="RFU", dimension="signal", factor=1.0, table=my_table)
```

`register_unit` without `table` edits the process-wide builtin map. The MCP
server copies that map when it loads settings, so call it before
`Settings.from_env` if you need the copy to see it.

Or point `EXCEL_UNITS_FILE` at a JSON file inside `EXCEL_ALLOWED_ROOTS`:

```json
{
  "units": [
    {
      "aliases": ["cups"],
      "canonical": "mL",
      "dimension": "volume",
      "factor": 236.588,
      "offset": 0,
      "note": "US customary cup"
    }
  ]
}
```

Aliases in the file replace builtin aliases with the same folded key.
Folding lowercases, drops spaces, maps the micro sign to `u`, maps the
degree sign to `deg`, and rewrites `mg.L-1` to `mg/l`.
