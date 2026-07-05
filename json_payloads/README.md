# json_payloads

Large deterministic JSON payloads plus a workbook performance harness.

## Files

- `generate_payloads.py`: regenerates every payload (seeded RNG, pure
  ASCII output so VBA's ANSI file read is lossless; unicode content is
  carried as `\uXXXX` escapes and exercises the parser's escape decoding).
- `Performance_Payloads_.bas`: the harness module (also imported in
  `ModernJsonInVBA.xlsm`).
- `tbl_*.json`: the payloads (generated, not committed):

| Payload | Rows | Shape |
|---|---:|---|
| `tbl_flat_10k.json` | 10,000 | 10 flat mixed-type columns (quick check) |
| `tbl_flat_100k.json` | 100,000 | same shape, mid-size |
| `tbl_flat_500k.json` | 500,000 | same shape, ~110 MB |
| `tbl_escapes_50k.json` | 50,000 | escape-dense strings, `\uXXXX`, surrogate pairs |
| `tbl_nested_50k.json` | 50,000 | nested objects (dotted columns) + tag arrays |
| `tbl_wide_5k_200c.json` | 5,000 | 200 columns (wide schema) |
| `tbl_numbers_200k.json` | 200,000 | numeric-heavy: big ints, floats, exponents |

## Running the suite

1. `python generate_payloads.py` (once, or after changing shapes)
2. Open `ModernJsonInVBA.xlsm`, press `ALT+F11`, open the Immediate window
   (`Ctrl+G`)
3. Run:

   ```
   Run_JsonPerfSuite
   ```

A markdown report prints to the Immediate window; paste it directly into a
GitHub issue or commit message. Each payload reports file read, `Json_Parse`
(model), `Json_Stringify`, table upsert create/refresh
(`Excel_UpsertListObjectFromJsonAtRoot`), and `Excel_ListObjectToJson`
export, with throughput per operation.

Payloads with `nested` in the name upsert with `nonTableArraysAsJson:=True`
so array columns are included as JSON text.
