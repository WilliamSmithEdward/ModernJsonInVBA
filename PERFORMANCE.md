# ModernJsonInVBA Performance

Real-world timings across the JSON parser and Excel ListObject surface.
Regenerate with `Run_JsonPerfMatrix` (json_payloads/Performance_Payloads_.bas).

- Generated: 2026-07-07 15:10:39
- Excel 16.0 (64-bit VBA)
- CPU: AMD Ryzen 7 9800X3D, 64 GB RAM
- Each cell is one wall-clock run (`Timer`); payloads are deterministic (seeded).

## Payloads

| Payload | Rows | Cols | Size (MB) |
|---|---:|---:|---:|
| flat_10k | 10,000 | 10 | 2.1 |
| nested_50k | 50,000 | 9 | 9.2 |
| escapes_50k | 50,000 | 6 | 14.8 |
| wide_5k_200c | 5,000 | 200 | 16.0 |
| flat_100k | 100,000 | 10 | 21.6 |
| numbers_200k | 200,000 | 8 | 25.9 |
| flat_500k | 500,000 | 10 | 109.5 |

## Timings (seconds)

| Step | flat_10k (2.1 MB) | nested_50k (9.2 MB) | escapes_50k (14.8 MB) | wide_5k_200c (16.0 MB) | flat_100k (21.6 MB) | numbers_200k (25.9 MB) | flat_500k (109.5 MB) |
|---|---:|---:|---:|---:|---:|---:|---:|
| Read file | 0.0273 | 0.0977 | 0.1602 | 0.1719 | 0.2266 | 0.2695 | 1.1367 |
| Json_Parse (JSON to model) | 0.2383 | 1.5547 | 1.3242 | 2.2383 | 2.3672 | 4.3867 | 11.8320 |
| Json_Stringify (model to JSON) | 0.2461 | 1.6563 | 1.3828 | 1.7227 | 2.5625 | 2.9727 | 12.8242 |
| Upsert create (JSON to ListObject) | 0.3164 | 2.4492 | 2.0625 | 2.8828 | 3.3906 | 4.6719 | 17.6680 |
| Upsert refresh | 0.3789 | 2.7070 | 2.3711 | 3.4805 | 3.9219 | 5.6563 | 20.1758 |
| Export (ListObject to JSON) | 0.1211 | 3.4766 | 0.9414 | 0.7500 | 1.2617 | 0.9297 | 6.4297 |
