# LIMIM_anon

A Python demonstration project that generates fictional infection-research data and creates an anonymised Excel workbook from it.

> **Data notice:** The generated records are synthetic and intended for software demonstration only. They are not real patient data or a clinically validated dataset. This project does not guarantee anonymisation or replace a formal data-protection review.

## What the project does

- Generates a workbook with 50 synthetic episodes and the selected infection-research fields.
- Replaces each NHS number with a keyed `STUDY-...` identifier.
- Applies a consistent, keyed date shift per patient to configured event dates, preserving that patient’s intervals. Shifted calendar dates cannot be compared across patients.
- Replaces date of birth with age at admission.
- Reduces full postcodes to outward areas and generalises locations.
- Pseudonymises orderer and prescriber IDs.
- Adds an `admission_year_group` based on the original admission year. The original year can appear in any position in the three-year group.

## Main files

- `simulate_data` — generates `synthetic_infection_research_selected_fields.xlsx`.
- `anonymize.py` — reads that workbook and writes `anonymized_infection_research.xlsx`.
- `anonymize_pseudocode.txt` — describes the anonymisation process.
- `generate_synthetic_data.py` — generates a separate, simpler synthetic workbook example.

## Requirements

Python 3 and these packages:

```powershell
python -m pip install pandas Faker openpyxl
```

## Run the current workflow

Run the following commands from the project folder.

### 1. Generate the synthetic input workbook

```powershell
python .\simulate_data
```

This creates `synthetic_infection_research_selected_fields.xlsx`.

### 2. Create the anonymised workbook

Close the input workbook in Excel before running the anonymizer. In PowerShell:

```powershell
$secureKey = Read-Host "Enter a secret key (at least 32 characters)" -AsSecureString
$ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
try {
    $env:ANONYMIZATION_KEY = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    python .\anonymize.py
}
finally {
    Remove-Item Env:ANONYMIZATION_KEY -ErrorAction SilentlyContinue
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    Remove-Variable secureKey, ptr -ErrorAction SilentlyContinue
}
```

This creates `anonymized_infection_research.xlsx` without overwriting the input.

Keep the key secret and store it securely. Reuse the same key only when you need consistent pseudonyms, date shifts, and admission-year-group choices. **Never commit the key to GitHub or store it in a workbook.**

## Important limitations

- The sample NHS numbers are fictional and intentionally fail NHS-number checksum validation. The output study IDs are not NHS numbers.
- Date shifting preserves intervals within each patient but changes the calendar dates; do not use shifted dates for comparisons between patients.
- `admission_year_group` is a broad range derived from the original admission year, not from the shifted date.
- Pseudonymisation and date shifting alone do not guarantee that real data is anonymous. Follow your organisation’s data-governance requirements before using or sharing real data.
