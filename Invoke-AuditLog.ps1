"""Print the Msid values found only in the QA CSV, in QA file order."""

import argparse
import csv
from pathlib import Path


def read_msids(path):
    with path.open(encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source, strict=True)
        if not reader.fieldnames or "Msid" not in reader.fieldnames:
            raise ValueError(f"{path}: missing column 'Msid'")
        for row in reader:
            msid = (row["Msid"] or "").strip()
            if msid:
                yield msid


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("qa", nargs="?", type=Path, default=Path("Qa_result.csv"))
    parser.add_argument("prod", nargs="?", type=Path, default=Path("Prod_result.csv"))
    args = parser.parse_args()

    try:
        seen = set(read_msids(args.prod))
        for msid in read_msids(args.qa):
            if msid not in seen:
                print(msid)
                seen.add(msid)
    except (OSError, ValueError, csv.Error) as error:
        parser.exit(1, f"{parser.prog}: {error}\n")


if __name__ == "__main__":
    main()
