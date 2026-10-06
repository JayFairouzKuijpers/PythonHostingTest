import argparse
import json
from pathlib import Path

from openpyxl import load_workbook


def read_model_output(
    excel_path: Path,
    sheet_name: str | None = None,
) -> list[dict[str, str]]:
    """Read the SOBnummer and Omschrijving columns from an Excel workbook."""
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet_name] if sheet_name else workbook.active
        rows = worksheet.iter_rows(values_only=True)
        header_row = next(rows, None)
        if header_row is None:
            raise ValueError(f"Worksheet '{worksheet.title}' is empty.")

        headers = {
            str(value).strip().casefold(): index
            for index, value in enumerate(header_row)
            if value is not None
        }
        if "SOBnummer" not in headers:
            raise ValueError("Missing required Excel column: SOBnummer")
        description_header = headers.get("omschrijving")
        if description_header is None:
            description_header = headers.get("object omschrijving")
        if description_header is None:
            raise ValueError(
                "Missing required Excel column: Omschrijving "
                "(or Object Omschrijving)."
            )

        number_index = headers["SOBnummer"]
        description_index = description_header
        records = []
        for row in rows:
            number = row[number_index] if number_index < len(row) else None
            description = (
                row[description_index]
                if description_index < len(row)
                else None
            )
            if number is None or description is None:
                continue

            number_text = str(number).strip()
            description_text = str(description).strip()
            if number_text and description_text:
                records.append(
                    {"SOBnummer": number_text, "Gestandaardiseerde Omschrijving": description_text}
                )
        return records
    finally:
        workbook.close()


def write_intermediate_json(
    records: list[dict[str, str]],
    output_path: Path,
) -> Path:
    """Store the extracted columns separately from the validation JSON."""
    output_path = Path(output_path)
    with output_path.open("w", encoding="utf-8") as file:
        json.dump(records, file, ensure_ascii=False, indent=2)
        file.write("\n")
    return output_path


def main() -> None:
    script_directory = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="Extract SOBnummer and Omschrijving columns to an intermediate JSON file."
    )
    parser.add_argument(
        "excel_file",
        nargs="?",
        type=Path,
        default=script_directory / "DummyDataOutput.xlsx",
        help="Source workbook (defaults to DummyDataOutput.xlsx).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=script_directory / "excel_objecten.json",
        help="Intermediate JSON destination (defaults to excel_objecten.json).",
    )
    parser.add_argument("--sheet", dest="sheet_name", help="Worksheet name to read.")
    args = parser.parse_args()

    records = read_model_output(args.excel_file, args.sheet_name)
    output_path = write_intermediate_json(records, args.output)
    print(f"Stored {len(records)} SOBnummer/Omschrijving pairs in {output_path}")


if __name__ == "__main__":
    main()