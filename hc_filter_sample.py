"""
hc_filter_sample.py

Simple filter/sort pass over a notification CSV export.

  1. Asks for the input CSV path.
  2. Drops any row where First Name or Last Name is blank, or is the
     placeholder text "[Unknown]" (case-insensitive).
  3. Sorts the remaining rows ascending by First Name, then Last Name,
     then Social Security Number.
  4. Writes the result to sample_10K.csv in the current directory.

Run:
    python hc_filter_sample.py
"""

import pandas as pd

COL_FIRST = "First Name"
COL_LAST = "Last Name"
COL_SSN = "Social Security Number"

OUTPUT_CSV = "sample_10K.csv"
UNKNOWN_PLACEHOLDER = "[unknown]"


def is_blank_or_unknown(value: str) -> bool:
    text = "" if value is None else str(value).strip()
    return text == "" or text.lower() == UNKNOWN_PLACEHOLDER


def main():
    input_path = input("Paste the input CSV path: ").strip().strip('"')

    df = pd.read_csv(input_path, dtype=str, keep_default_na=False)

    mask = ~(df[COL_FIRST].apply(is_blank_or_unknown) | df[COL_LAST].apply(is_blank_or_unknown))
    filtered = df[mask]

    sorted_df = filtered.sort_values(by=[COL_FIRST, COL_LAST, COL_SSN], ascending=True)

    sorted_df.to_csv(OUTPUT_CSV, index=False)
    print(f"Read {len(df)} rows, kept {len(sorted_df)} after filtering, wrote {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
