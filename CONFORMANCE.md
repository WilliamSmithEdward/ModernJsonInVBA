# ModernJsonInVBA Conformance

Results of running the parser against
[JSONTestSuite](https://github.com/nst/JSONTestSuite), the standard RFC 8259
conformance corpus (318 files).

- Run: 2026-07-07, library v3.7.0, Excel 16.0 (64-bit VBA)
- Parser entry point: `Json_TryParse` (same grammar as `Json_Parse`)

## Results

| Category | Meaning | Result |
|---|---|---|
| `y_` (95 files) | valid JSON, must accept | **95/95 accepted** |
| `n_` (188 files) | invalid JSON, must reject | **176/176 rejected**, 12 byte-level files exercised at the decoding layer (below) |
| `i_` (35 files) | implementation-defined | 22 run, choices listed below; 13 byte-level |

No test crashed the host. The deep-nesting attacks
(`n_structure_100000_opening_arrays` and friends) reject through VBA's
trappable "Out of stack space" error rather than terminating the process.

## Layering: bytes vs text

The parser's input is a VBA `String` (UTF-16), so the corpus files whose
point is invalid *bytes* (malformed UTF-8 sequences, UTF-16 without a BOM,
Latin-1 bytes) do not reach the parser as-is; they exercise the decoding
layer instead. That layer is `Json_ReadTextFile`, which the `*FileToJson`
adapters use:

- a strict pure-VBA UTF-8 decoder rejects overlong encodings, surrogate
  code points, code points past U+10FFFF, stray continuation bytes, and
  truncated sequences,
- files that fail strict UTF-8 validation fall back to the system ANSI
  codepage (the legacy behavior for old exports),
- UTF-8 / UTF-16 LE / UTF-16 BE byte-order marks are detected and stripped.

The 25 corpus files in this class (12 `n_`, 13 `i_`) are reported separately
rather than counted as parser accepts or rejects.

## Implementation-defined choices (`i_` files)

| Case | Choice | Mechanism |
|---|---|---|
| Numbers beyond Double range (`1e308+`, huge exponents) | reject | `CDbl` overflow raises, surfaced as parse failure |
| Integers beyond Long range (`123123e100000` style big ints) | accept as `Double` | standard precision-loss behavior |
| Real underflow (`1e-2000`) | accept as `0` | `CDbl` underflow to zero |
| Lone or invalid `\uXXXX` surrogate escapes | reject (error 527) | surrogate pairs are validated strictly |
| Nesting deeper than the VBA call stack allows | reject | trappable "Out of stack space"; 500 nested arrays reject, typical documents (under ~250 levels) parse |
| U+FEFF (BOM) as the first character of *string* input | reject (error 701) | file-level BOMs are stripped by `Json_ReadTextFile` before the parser runs; text callers own their own input |

Every choice above is conforming: RFC 8259 leaves these cases to the
implementation.

## Reproducing

1. Clone [JSONTestSuite](https://github.com/nst/JSONTestSuite).
2. For each file in `test_parsing/`, decode the bytes (strict UTF-8; set
   aside files that fail as decoding-layer cases).
3. Feed each decoded string to `Json_TryParse` and compare the Boolean
   against the file's `y_`/`n_` prefix.
