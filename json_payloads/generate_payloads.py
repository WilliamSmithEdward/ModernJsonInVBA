"""Generate deterministic large JSON payloads for the workbook perf suite
(Performance_Payloads_.bas / Run_JsonPerfSuite).

All output is pure ASCII (ensure_ascii=True), so VBA's ANSI file read is
lossless and unicode content still exercises the parser's \\uXXXX decoding.
Re-running always produces identical files (seeded RNG).

Usage:  python generate_payloads.py
"""
import json
import os
import random

HERE = os.path.dirname(os.path.abspath(__file__))

CITIES = ["Springfield", "Riverton", "Lakewood", "Fairview", "Georgetown",
          "Ashland", "Milton", "Clayton", "Dayton", "Salem"]
REGIONS = ["NE", "NW", "SE", "SW", "C"]
STATUSES = ["new", "active", "on-hold", "closed"]

COMPACT = (",", ":")


def write_array(path, count, make_obj):
    with open(path, "w", encoding="ascii", newline="") as f:
        f.write("[")
        for i in range(count):
            if i:
                f.write(",")
            f.write(json.dumps(make_obj(i), separators=COMPACT, ensure_ascii=True))
        f.write("]")
    size = os.path.getsize(path)
    print(f"{os.path.basename(path):28s} {count:>9,} rows  {size / 1024 / 1024:8.1f} MB")


def flat_row(i):
    rnd = random.Random(i)  # per-row determinism
    return {
        "id": i + 1,
        "name": f"user_{i + 1}",
        "email": f"user{i + 1}@example.com",
        "active": (i % 2 == 0),
        "score": round(rnd.uniform(0, 100), 4),
        "amount": rnd.randint(1, 500000),
        "city": CITIES[i % len(CITIES)],
        "region": REGIONS[i % len(REGIONS)],
        "joined": f"20{10 + i % 15:02d}-{1 + i % 12:02d}-{1 + i % 28:02d}",
        "note": f"plain note text for row {i + 1} with no escapes at all",
    }


def escapes_row(i):
    emoji = "\U0001F600\U0001F680"  # surrogate pairs when ASCII-escaped
    return {
        "id": i + 1,
        "quoted": f'she said "hello" and \\ backslash {i}',
        "multiline": f"line one {i}\nline two\ttabbed\r\nline three",
        "unicode": f"café über straße {emoji} row {i}",
        "path": f"C:\\data\\files\\row_{i}\\payload.json",
        "control": f"bell and null-ish markers {i}",
    }


def nested_row(i):
    rnd = random.Random(i * 7 + 3)
    return {
        "id": i + 1,
        "status": STATUSES[i % len(STATUSES)],
        "customer": {
            "name": f"customer_{i + 1}",
            "tier": 1 + i % 4,
            "address": {
                "city": CITIES[i % len(CITIES)],
                "zip": f"{10000 + i % 89999:05d}",
            },
        },
        "metrics": {
            "score": round(rnd.uniform(0, 10), 3),
            "visits": rnd.randint(0, 5000),
        },
        "tags": [f"tag_{i % 23}", f"tag_{i % 7}", f"grp_{i % 41}"],
    }


def wide_row(i):
    row = {"id": i + 1}
    for c in range(1, 200):
        if c % 3 == 0:
            row[f"col_{c:03d}"] = (i * c) % 10000
        elif c % 3 == 1:
            row[f"col_{c:03d}"] = f"v{(i + c) % 997}"
        else:
            row[f"col_{c:03d}"] = round(((i + 1) * c % 7919) / 13.0, 4)
    return row


def numbers_row(i):
    rnd = random.Random(i * 13 + 1)
    return {
        "id": i + 1,
        "big": rnd.randint(10**12, 10**15),         # beyond Long -> Double
        "neg": -rnd.randint(1, 10**9),
        "float6": round(rnd.uniform(-1e6, 1e6), 6),
        "exp": float(f"{rnd.uniform(1, 9):.6f}e{rnd.randint(-20, 20)}"),
        "small": rnd.randint(0, 9),
        "zero": 0,
        "ratio": round(rnd.random(), 12),
    }


def main():
    random.seed(42)
    print("generating payloads into", HERE)

    write_array(os.path.join(HERE, "tbl_flat_10k.json"), 10_000, flat_row)
    write_array(os.path.join(HERE, "tbl_flat_100k.json"), 100_000, flat_row)
    write_array(os.path.join(HERE, "tbl_flat_500k.json"), 500_000, flat_row)
    write_array(os.path.join(HERE, "tbl_escapes_50k.json"), 50_000, escapes_row)
    write_array(os.path.join(HERE, "tbl_nested_50k.json"), 50_000, nested_row)
    write_array(os.path.join(HERE, "tbl_wide_5k_200c.json"), 5_000, wide_row)
    write_array(os.path.join(HERE, "tbl_numbers_200k.json"), 200_000, numbers_row)

    print("done")


if __name__ == "__main__":
    main()
