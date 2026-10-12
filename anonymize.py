"""Anonymise the synthetic infection-research workbook.

Edit the settings below to match the rules for your data. The input workbook is
never overwritten; the result is saved as anonymized_infection_research.xlsx.
"""

import hashlib
import hmac
import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.table import TableColumn


# ---------- Settings you may need to change ----------

INPUT_FILE = Path(__file__).with_name("synthetic_infection_research_selected_fields.xlsx")
OUTPUT_FILE = Path(__file__).with_name("anonymized_infection_research.xlsx")

# Replace ward names with broad care groups and hospital names with "Hospital".
WARD_GROUPS = {
    "Acute": "Acute care",
    "Critical Care": "ITU",
    "Respiratory": "Medicine",
    "Surgical": "Surgery",
    "Emergency": "Emergency care",
}

DATE_COLUMNS = {
    "microbiology_datetime_collected",
    "microbiology_datetime_reported",
    "microbiology_datetime_stored",
    "collected_datetime",
    "reported_datetime",
    "start_datetime",
    "end_datetime",
    "administration_datetime",
    "admission_datetime",
    "transfer_datetime",
    "date_of_death",
    "radiology_study_datetime",
    "observation_datetime",
    "clinical_encounter_datetime",
}
LINKED_CODE_COLUMNS = {
    "microbiology_orderer_id",
    "haematology_orderer_id",
    "prescriber_id",
}
LOCATION_COLUMNS = {
    "haematology_specimen_location",
    "prescription_location",
    "admission_location",
    "transfer_location",
    "observation_location",
}
REQUIRED_COLUMNS = {
    "microbiology_datetime_collected",
    "microbiology_datetime_reported",
    "microbiology_datetime_stored",
    "microbiology_orderer_id",
    "collected_datetime",
    "reported_datetime",
    "haematology_specimen_location",
    "haematology_orderer_id",
    "prescription_location",
    "prescriber_id",
    "start_datetime",
    "end_datetime",
    "administration_datetime",
    "nhs_number",
    "postcode",
    "admission_datetime",
    "admission_location",
    "transfer_datetime",
    "transfer_location",
    "date_of_birth",
    "date_of_death",
    "radiology_study_datetime",
    "observation_datetime",
    "observation_location",
    "clinical_encounter_datetime",
}


def token(value, purpose, secret):
    """Turn an identifier into a repeatable, unreadable code."""
    message = f"{purpose}|{value}".encode("utf-8")
    return hmac.new(secret, message, hashlib.sha256).hexdigest()[:24]


def patient_date_shift(patient_id, secret):
    """Give each patient one repeatable random shift between 1 and 50,000 days."""
    message = f"date|{patient_id}".encode("utf-8")
    limit = (1 << 256) - ((1 << 256) % 50_000)
    attempt = 0
    while True:
        digest = hmac.new(
            secret, message + attempt.to_bytes(4, "big"), hashlib.sha256
        ).digest()
        number = int.from_bytes(digest, "big")
        if number < limit:  # Ignore the tiny uneven tail so every day is equally likely.
            return number % 50_000 + 1
        attempt += 1


def patient_admission_year_group(admission, patient_id, secret):
    """Make a stable keyed-random three-year group containing the original year."""
    message = f"admission-year-group|{patient_id}|{admission.year}".encode("utf-8")
    position = int.from_bytes(
        hmac.new(secret, message, hashlib.sha256).digest(), "big"
    ) % 3
    first_year = admission.year - position
    return f"{first_year}-{first_year + 2}"


def as_datetime(value, row_number, column):
    """Read a date from Excel or from a date string."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, str):
        # Some Excel exports use typographic dashes in otherwise standard dates.
        value = value.strip().translate(str.maketrans("‐‑‒–−", "-----"))
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError as error:
            raise ValueError(
                f"Cannot read {column!r} as a date on spreadsheet row {row_number}."
            ) from error
    raise ValueError(
        f"Unexpected value in {column!r} on spreadsheet row {row_number}."
    )


def age_on_admission(birth_date, admission_date, row_number):
    """Replace date of birth with age in years; cap ages at 100."""
    born = as_datetime(birth_date, row_number, "date_of_birth").date()
    admitted = admission_date.date()
    age = admitted.year - born.year
    if (admitted.month, admitted.day) < (born.month, born.day):
        age -= 1
    if age < 0:
        raise ValueError(f"Date of birth is after admission on row {row_number}.")
    return "≥100" if age >= 100 else age


def postcode_area(postcode, row_number):
    """Keep only the outward part of a UK postcode, such as L1."""
    if postcode is None or not str(postcode).strip():
        return postcode
    match = re.fullmatch(
        r"([A-Z]{1,2}\d[A-Z\d]?)\s*\d[A-Z]{2}",
        re.sub(r"\s+", "", str(postcode).upper()),
    )
    if not match:
        raise ValueError(f"Cannot recognise postcode on spreadsheet row {row_number}.")
    return match.group(1)


def generalise_location(location, column):
    """Remove specific facility or ward names while retaining broad location type."""
    if location is None or not str(location).strip():
        return location
    if column == "admission_location":
        return "Hospital"

    normalise = lambda text: re.sub(
        r"\s+", " ", text.strip().translate(str.maketrans("‐‑‒–−", "-----"))
    ).casefold()
    name = normalise(str(location))

    matches = {
        group for term, group in WARD_GROUPS.items()
        if normalise(term) in name
    }
    if len(matches) == 1:
        return matches.pop()
    if len(matches) > 1:
        raise ValueError(
            f"Location {location!r} contains terms assigned to different groups. "
            "Review WARD_GROUPS."
        )
    raise ValueError(
        f"No broad group is set for location {location!r} in {column!r}. "
        "Add its term to WARD_GROUPS."
    )


def sync_table_headers(sheet):
    """Keep Excel Table column labels aligned with worksheet headers."""
    for table_name in list(sheet.tables):
        table = sheet.tables[table_name]
        min_col, min_row, _, _ = range_boundaries(table.ref)
        for index, column in enumerate(table.tableColumns, start=1):
            column.id = index
            column.name = str(sheet.cell(min_row, min_col + index - 1).value)
        if table.autoFilter is not None:
            table.autoFilter.ref = table.ref


def extend_tables_for_column(sheet, column_index, header):
    """Include a newly added worksheet column in tables that cover the data."""
    for table_name in list(sheet.tables):
        table = sheet.tables[table_name]
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if min_row == 1 and max_col == column_index - 1:
            table.ref = (
                f"{get_column_letter(min_col)}{min_row}:"
                f"{get_column_letter(column_index)}{max_row}"
            )
            if table.autoFilter is not None:
                table.autoFilter.ref = table.ref
            table.tableColumns.append(
                TableColumn(id=len(table.tableColumns) + 1, name=header)
            )


def main():
    # Reuse the secret only when consistent pseudonyms and date shifts are needed.
    secret_text = os.environ.get("ANONYMIZATION_KEY")
    if not secret_text or len(secret_text) < 32:
        raise ValueError(
            "Set ANONYMIZATION_KEY to a secret of at least 32 characters. "
            "Keep it separate from the workbook."
        )
    secret = secret_text.encode("utf-8")

    workbook = load_workbook(INPUT_FILE)
    if len(workbook.worksheets) != 1:
        raise ValueError("This starter script expects one worksheet.")
    sheet = workbook.active
    headers = [cell.value for cell in sheet[1]]
    if len(headers) != len(set(headers)) or any(header is None for header in headers):
        raise ValueError("The first row must contain unique, non-empty column names.")

    columns = {name: index + 1 for index, name in enumerate(headers)}
    missing = REQUIRED_COLUMNS - columns.keys()
    unknown = columns.keys() - REQUIRED_COLUMNS
    if missing or unknown:
        raise ValueError(
            f"Missing required columns: {sorted(missing)}; "
            f"unexpected columns: {sorted(unknown)}."
        )

    year_group_column = sheet.max_column + 1
    sheet.cell(1, year_group_column).value = "admission_year_group"

    for row in range(2, sheet.max_row + 1):
        nhs_number = sheet.cell(row, columns["nhs_number"]).value
        if nhs_number is None or not str(nhs_number).strip():
            raise ValueError(f"Missing NHS number on spreadsheet row {row}.")

        admission_cell = sheet.cell(row, columns["admission_datetime"])
        admission = as_datetime(admission_cell.value, row, "admission_datetime")
        sheet.cell(row, year_group_column).value = (
            patient_admission_year_group(admission, nhs_number, secret)
        )

        # Replace the exact birth date with age at admission.
        dob_cell = sheet.cell(row, columns["date_of_birth"])
        if dob_cell.value is not None:
            dob_cell.value = age_on_admission(dob_cell.value, admission, row)
            dob_cell.number_format = "General"

        # Apply the same patient-specific shift to every activity and death date.
        days = patient_date_shift(nhs_number, secret)
        for name in DATE_COLUMNS:
            cell = sheet.cell(row, columns[name])
            if cell.value is not None:
                original = cell.value
                shifted = as_datetime(original, row, name) + timedelta(days=days)
                if isinstance(original, str):
                    cell.value = shifted.isoformat(sep=" " if " " in original else "T")
                elif isinstance(original, date) and not isinstance(original, datetime):
                    cell.value = shifted.date()
                else:
                    cell.value = shifted

        # Replace locations, postcode detail, and directly identifying numbers.
        for name in LOCATION_COLUMNS:
            cell = sheet.cell(row, columns[name])
            cell.value = generalise_location(cell.value, name)

        postcode_cell = sheet.cell(row, columns["postcode"])
        postcode_cell.value = postcode_area(postcode_cell.value, row)
        sheet.cell(row, columns["nhs_number"]).value = "STUDY-" + token(
            nhs_number, "nhs_number", secret
        )
        for name in LINKED_CODE_COLUMNS:
            cell = sheet.cell(row, columns[name])
            if cell.value is not None:
                cell.value = token(cell.value, name, secret)

    # Label transformed values without changing the selected field set.
    sheet.cell(1, columns["postcode"]).value = "postcode_area"
    sheet.cell(1, columns["date_of_birth"]).value = "age_at_admission"
    extend_tables_for_column(sheet, year_group_column, "admission_year_group")
    sync_table_headers(sheet)

    workbook.save(OUTPUT_FILE)
    print(f"Done. Anonymised workbook saved as {OUTPUT_FILE.name}")


if __name__ == "__main__":
    main()
