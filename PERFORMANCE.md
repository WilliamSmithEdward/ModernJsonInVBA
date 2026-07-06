# ModernJsonInVBA Performance

Real-world timings across the JSON parser and Excel ListObject surface.
Regenerate with `Run_JsonPerfMatrix` (json_payloads/Performance_Payloads_.bas).

- Generated: 2026-07-05 17:23:27
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

| Step | flat_10k | nested_50k | escapes_50k | wide_5k_200c | flat_100k | numbers_200k | flat_500k |
|---|---:|---:|---:|---:|---:|---:|---:|
| Read file | 0.0273 | 0.0938 | 0.1523 | 0.1641 | 0.2227 | 0.2695 | 1.1289 |
| Json_Parse (JSON to model) | 0.2695 | 1.5586 | 1.3203 | 2.2344 | 2.3672 | 4.3984 | 11.8789 |
| Json_Stringify (model to JSON) | 0.2695 | 1.6719 | 1.4023 | 1.7266 | 2.5742 | 3.0430 | 12.9492 |
| Upsert create (JSON to ListObject) | 0.3281 | 2.4727 | 2.0625 | 2.8828 | 3.4727 | 4.7383 | 18.2656 |
| Upsert refresh | 0.3789 | 2.7422 | 2.3711 | 3.4844 | 3.9688 | 5.7305 | 20.3828 |
| Export (ListObject to JSON) | 0.1328 | 3.5195 | 0.9609 | 0.8477 | 1.3945 | 1.1484 | 7.1094 |
