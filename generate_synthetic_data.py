"""Generate a fictional workbook for demonstrating the anonymisation script."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


OUTPUT_FILE = Path(__file__).with_name("synthetic_pre_anon_table.xlsx")

HEADERS = (
    "patient_id",
    "hospital_name",
    "ward_name",
    "admission_datetime",
    "postcode",
    "date_of_birth",
)

ROWS = (
    ("123 456 7891", "Synthetic General North", "Acute Medical Unit", "2018-03-12 14:22", "L1 1ZZ", "1973-07-15"),
    ("123 456 7891", "Synthetic General North", "Critical Care Unit", "2018-03-13 09:10", "L1 1ZZ", "1973-07-15"),
    ("231 546 7981", "Synthetic General South", "Respiratory Assessment Unit", "2017-11-01 21:40", "CH1 1ZZ", "1939-12-01"),
    ("231 546 7981", "Synthetic General South", "Surgical Assessment Unit", "2018-11-02 00:10", "CH1 1ZZ", "1939-12-01"),
    ("321 654 9870", "Synthetic Community Hospital", "Cardiology Day Unit", "2019-04-05 10:00", "L14 3LB", "1917-11-01"),
    ("321 654 9870", "Synthetic Community Hospital", "Theatre Suite", "2020-04-05 13:00", "L14 3LB", "1917-11-01"),
    ("900 000 0004", "Synthetic General North", "Acute Medical Unit", "2019-02-10 08:15", "L1 1ZZ", "1988-03-02"),
    ("900 000 0004", "Synthetic General North", "Surgical Ward 4", "2019-02-13 16:20", "L1 1ZZ", "1988-03-02"),
    ("900 000 0004", "Synthetic General South", "Respiratory Assessment Unit", "2020-06-01 11:05", "CH1 1ZZ", "1988-03-02"),
    ("900 000 0005", "Synthetic General North", "Day Unit", "2018-03-12 09:00", "M1 1ZZ", "1954-09-19"),
    ("900 000 0005", "Synthetic General North", "Theatre Suite", "2018-06-22 14:45", "M1 1ZZ", "1954-09-19"),
    ("900 000 0006", "Synthetic General South", "Critical Care Unit", "2020-04-19 23:30", "LS1 1ZZ", "1916-04-20"),
    ("900 000 0006", "Synthetic General South", "Acute Medical Unit", "2020-04-21 07:10", "LS1 1ZZ", "1916-04-20"),
    ("900 000 0007", "Synthetic General North", "Respiratory Assessment Unit", "2017-12-10 13:25", "B1 1ZZ", "2002-12-14"),
    ("900 000 0007", "Synthetic General North", "Surgical Ward 4", "2018-01-05 10:40", "B1 1ZZ", "2002-12-14"),
    ("900 000 0008", "Synthetic General South", "Cardiology Day Unit", "2021-07-08 12:00", "WA1 1ZZ", "1977-05-30"),
)


def generate_workbook(output_file: Path = OUTPUT_FILE) -> Path:
    """Write the fictional example data to a separate Excel workbook."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Synthetic data"
    worksheet.append(HEADERS)
    for row in ROWS:
        worksheet.append(row)

    table = Table(displayName="SyntheticPatientData", ref=worksheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)
    worksheet.freeze_panes = "A2"
    workbook.properties.title = "Synthetic patient data for anonymisation demonstration"
    workbook.properties.subject = "Entirely fictional example data"
    workbook.properties.description = (
        "All records, identifiers, dates, and locations are fictional."
    )
    workbook.save(output_file)
    return output_file


if __name__ == "__main__":
    print(f"Created fictional example workbook: {generate_workbook().name}")
