"""Convierte archivos XLS y XLSX a CSV."""

from pathlib import Path

import pandas as pd


def convert_excel_to_csv(input_file, output_file=None):
    """Convert an XLS or XLSX file to CSV and return the output path."""
    input_path = Path(input_file)
    suffix = input_path.suffix.casefold()
    if suffix not in {".xls", ".xlsx"}:
        raise ValueError(f"Unsupported Excel file extension: {input_path.suffix}")

    output_path = (
        Path(output_file) if output_file else input_path.with_suffix(".csv")
    )

    engine = "xlrd" if suffix == ".xls" else "openpyxl"
    # dtype=str preserves leading zeroes in plates, codes, and ticket numbers.
    dataframe = pd.read_excel(input_path, dtype=str, engine=engine)
    # utf-8-sig: Excel muestra bien las tildes; load_records lo lee sin problema.
    dataframe.to_csv(output_path, index=False, encoding="utf-8-sig")
    return output_path


def main():
    excel_files = sorted(
        path for path in Path.cwd().iterdir()
        if path.is_file() and path.suffix.casefold() in {".xls", ".xlsx"}
    )
    if not excel_files:
        print("No XLS or XLSX file found in the current directory.")
        return

    input_path = excel_files[0]
    try:
        output_path = convert_excel_to_csv(input_path)
    except (OSError, ImportError, ValueError) as error:
        print(f"Error converting file: {error}")
        return

    print(f"Successfully converted '{input_path}' to '{output_path}'.")


if __name__ == "__main__":
    main()