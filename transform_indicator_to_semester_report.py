import argparse
from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook, load_workbook


def to_number(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return 0.0
        text = text.replace(",", "")
        try:
            return float(text)
        except ValueError:
            return 0.0
    return 0.0


def get_header_index(headers: List[str], names: List[str]) -> dict:
    normalized = {name.strip().lower(): idx for idx, name in enumerate(headers)}
    result = {}
    for name in names:
        result[name] = normalized.get(name.strip().lower())
    return result


def build_semester_report(input_path: Path, output_path: Path, source_sheet: str = "Indicator", target_sheet: str = "Semester_Report") -> Path:
    wb = load_workbook(input_path, data_only=False)
    if source_sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{source_sheet}' not found in {input_path}")

    source_ws = wb[source_sheet]
    headers = [cell.value if cell.value is not None else "" for cell in next(source_ws.iter_rows(min_row=1, max_row=1))]
    header_map = get_header_index(headers, [
        "Year",
        "Organization",
        "Organization ",
        "Project Name",
        "indicator",
        "Q1 Target",
        "Q1 U1 Male",
        "Q1 U1 Female",
        "Q1 1-5 Male ",
        "Q1 1-5 Female",
        "Q1 Total",
        "Q2 Target",
        "Q2 U1 Male",
        "Q2 U1 Female",
        "Q2 1-5 Male ",
        "Q2 1-5 Female",
        "Q2 Total",
        "Q3 Target",
        "Q3 U1 Male",
        "Q3 U1 Female",
        "Q3 1-5 Male ",
        "Q3 1-5 Female",
        "Q3 Total",
        "Q4 Target",
        "Q4 U1 Male",
        "Q4 U1 Female",
        "Q4 1-5 Male ",
        "Q4 1-5 Female",
        "Q4 Total",
    ])

    required_headers = [
        "Year",
        "Organization",
        "Project Name",
        "indicator",
        "Q1 Target",
        "Q1 U1 Male",
        "Q1 U1 Female",
        "Q1 1-5 Male ",
        "Q1 1-5 Female",
        "Q1 Total",
        "Q2 Target",
        "Q2 U1 Male",
        "Q2 U1 Female",
        "Q2 1-5 Male ",
        "Q2 1-5 Female",
        "Q2 Total",
        "Q3 Target",
        "Q3 U1 Male",
        "Q3 U1 Female",
        "Q3 1-5 Male ",
        "Q3 1-5 Female",
        "Q3 Total",
        "Q4 Target",
        "Q4 U1 Male",
        "Q4 U1 Female",
        "Q4 1-5 Male ",
        "Q4 1-5 Female",
        "Q4 Total",
    ]
    for name in required_headers:
        if header_map.get(name) is None:
            raise ValueError(f"Required column '{name}' was not found in sheet '{source_sheet}'")

    if target_sheet in wb.sheetnames:
        del wb[target_sheet]

    target_ws = wb.create_sheet(title=target_sheet)
    target_ws.append([
        "Year",
        "Organization",
        "Project Name",
        "indicator",
        "S1 Target",
        "S1 Male",
        "S1 Female",
        "S2 Target",
        "S2 Male",
        "S2 Female",
        "S2 Total",
        "Annual Target",
        "Annual Male",
        "Annual Female",
        "Annual Total",
    ])

    for row in source_ws.iter_rows(min_row=2, values_only=True):
        if not any(cell is not None and str(cell).strip() != "" for cell in row):
            continue

        year = row[header_map["Year"]] if header_map["Year"] is not None and header_map["Year"] < len(row) else ""
        organization = row[header_map["Organization"]] if header_map["Organization"] is not None and header_map["Organization"] < len(row) else None
        if organization is None and header_map.get("Organization ") is not None and header_map["Organization "] < len(row):
            organization = row[header_map["Organization "]]
        project_name = row[header_map["Project Name"]] if header_map["Project Name"] is not None and header_map["Project Name"] < len(row) else ""
        indicator = row[header_map["indicator"]] if header_map["indicator"] is not None and header_map["indicator"] < len(row) else ""

        q1_target = to_number(row[header_map["Q1 Target"]] if header_map["Q1 Target"] is not None and header_map["Q1 Target"] < len(row) else None)
        q1_u1_male = to_number(row[header_map["Q1 U1 Male"]] if header_map["Q1 U1 Male"] is not None and header_map["Q1 U1 Male"] < len(row) else None)
        q1_u1_female = to_number(row[header_map["Q1 U1 Female"]] if header_map["Q1 U1 Female"] is not None and header_map["Q1 U1 Female"] < len(row) else None)
        q1_1_5_male = to_number(row[header_map["Q1 1-5 Male "]] if header_map["Q1 1-5 Male "] is not None and header_map["Q1 1-5 Male "] < len(row) else None)
        q1_1_5_female = to_number(row[header_map["Q1 1-5 Female"]] if header_map["Q1 1-5 Female"] is not None and header_map["Q1 1-5 Female"] < len(row) else None)
        q1_total = to_number(row[header_map["Q1 Total"]] if header_map["Q1 Total"] is not None and header_map["Q1 Total"] < len(row) else None)

        q2_target = to_number(row[header_map["Q2 Target"]] if header_map["Q2 Target"] is not None and header_map["Q2 Target"] < len(row) else None)
        q2_u1_male = to_number(row[header_map["Q2 U1 Male"]] if header_map["Q2 U1 Male"] is not None and header_map["Q2 U1 Male"] < len(row) else None)
        q2_u1_female = to_number(row[header_map["Q2 U1 Female"]] if header_map["Q2 U1 Female"] is not None and header_map["Q2 U1 Female"] < len(row) else None)
        q2_1_5_male = to_number(row[header_map["Q2 1-5 Male "]] if header_map["Q2 1-5 Male "] is not None and header_map["Q2 1-5 Male "] < len(row) else None)
        q2_1_5_female = to_number(row[header_map["Q2 1-5 Female"]] if header_map["Q2 1-5 Female"] is not None and header_map["Q2 1-5 Female"] < len(row) else None)
        q2_total = to_number(row[header_map["Q2 Total"]] if header_map["Q2 Total"] is not None and header_map["Q2 Total"] < len(row) else None)

        q3_target = to_number(row[header_map["Q3 Target"]] if header_map["Q3 Target"] is not None and header_map["Q3 Target"] < len(row) else None)
        q3_u1_male = to_number(row[header_map["Q3 U1 Male"]] if header_map["Q3 U1 Male"] is not None and header_map["Q3 U1 Male"] < len(row) else None)
        q3_u1_female = to_number(row[header_map["Q3 U1 Female"]] if header_map["Q3 U1 Female"] is not None and header_map["Q3 U1 Female"] < len(row) else None)
        q3_1_5_male = to_number(row[header_map["Q3 1-5 Male "]] if header_map["Q3 1-5 Male "] is not None and header_map["Q3 1-5 Male "] < len(row) else None)
        q3_1_5_female = to_number(row[header_map["Q3 1-5 Female"]] if header_map["Q3 1-5 Female"] is not None and header_map["Q3 1-5 Female"] < len(row) else None)
        q3_total = to_number(row[header_map["Q3 Total"]] if header_map["Q3 Total"] is not None and header_map["Q3 Total"] < len(row) else None)

        q4_target = to_number(row[header_map["Q4 Target"]] if header_map["Q4 Target"] is not None and header_map["Q4 Target"] < len(row) else None)
        q4_u1_male = to_number(row[header_map["Q4 U1 Male"]] if header_map["Q4 U1 Male"] is not None and header_map["Q4 U1 Male"] < len(row) else None)
        q4_u1_female = to_number(row[header_map["Q4 U1 Female"]] if header_map["Q4 U1 Female"] is not None and header_map["Q4 U1 Female"] < len(row) else None)
        q4_1_5_male = to_number(row[header_map["Q4 1-5 Male "]] if header_map["Q4 1-5 Male "] is not None and header_map["Q4 1-5 Male "] < len(row) else None)
        q4_1_5_female = to_number(row[header_map["Q4 1-5 Female"]] if header_map["Q4 1-5 Female"] is not None and header_map["Q4 1-5 Female"] < len(row) else None)
        q4_total = to_number(row[header_map["Q4 Total"]] if header_map["Q4 Total"] is not None and header_map["Q4 Total"] < len(row) else None)

        s1_target = q1_target + q2_target
        s1_male = q1_u1_male + q1_1_5_male + q2_u1_male + q2_1_5_male
        s1_female = q1_u1_female + q1_1_5_female + q2_u1_female + q2_1_5_female
        s2_target = q3_target + q4_target
        s2_male = q3_u1_male + q3_1_5_male + q4_u1_male + q4_1_5_male
        s2_female = q3_u1_female + q3_1_5_female + q4_u1_female + q4_1_5_female
        s2_total = q3_total + q4_total
        annual_target = q1_target + q2_target + q3_target + q4_target
        annual_male = q1_u1_male + q1_1_5_male + q2_u1_male + q2_1_5_male + q3_u1_male + q3_1_5_male + q4_u1_male + q4_1_5_male
        annual_female = q1_u1_female + q1_1_5_female + q2_u1_female + q2_1_5_female + q3_u1_female + q3_1_5_female + q4_u1_female + q4_1_5_female
        annual_total = q1_total + q2_total + q3_total + q4_total

        target_ws.append([
            year,
            organization,
            project_name,
            indicator,
            s1_target,
            s1_male,
            s1_female,
            s2_target,
            s2_male,
            s2_female,
            s2_total,
            annual_target,
            annual_male,
            annual_female,
            annual_total,
        ])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Transform the Indicator sheet into a semester report sheet in an Excel workbook")
    parser.add_argument("input_workbook", nargs="?", default=r"C:\Users\mssam\OneDrive\Data automation\KDHW\quarterly compile.update.xlsx")
    parser.add_argument("output_workbook", nargs="?", default=None)
    parser.add_argument("--source-sheet", default="Indicator")
    parser.add_argument("--target-sheet", default="Semester_Report")
    args = parser.parse_args()

    input_path = Path(args.input_workbook)
    if args.output_workbook:
        output_path = Path(args.output_workbook)
    else:
        output_path = input_path.with_name(input_path.stem + "_semester_report.xlsx")

    result = build_semester_report(input_path=input_path, output_path=output_path, source_sheet=args.source_sheet, target_sheet=args.target_sheet)
    print(f"Created {result}")


if __name__ == "__main__":
    main()
