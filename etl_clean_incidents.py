import pandas as pd
import re
from pathlib import Path

RAW_INCIDENT_FILE = "data/raw/Safety incident-accident report (1).xlsx"
RAW_FAS_FILE = "data/raw/FAS & FHS Report all WH.xlsx"
OUT_DIR = Path("data/processed")

OUT_DIR.mkdir(parents=True, exist_ok=True)

# ----------------------------
# Helper functions
# ----------------------------
def extract_date(text):
    if not isinstance(text, str):
        return None
    match = re.search(r"\d{2}[-/]\d{2}[-/]\d{4}", text)
    if match:
        return pd.to_datetime(match.group(), dayfirst=True, errors="coerce")
    return None

def classify_case(text):
    if not isinstance(text, str):
        return "Unclassified"
    t = text.lower()

    if "near miss" in t:
        return "Near Miss"
    if "first aid" in t:
        return "Incident"
    if "medical treatment" in t:
        return "Accident"
    if "lost time" in t:
        return "LTI"
    if "accident" in t:
        return "Accident"
    if "incident" in t:
        return "Incident"

    return "Unclassified"

def extract_field(label, text):
    if not isinstance(text, str):
        return None
    pattern = rf"{label}\s*:\s*(.+)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else None

# ----------------------------
# LOAD INCIDENT FILE
# ----------------------------
xls = pd.ExcelFile(RAW_INCIDENT_FILE)
records = []
incident_counter = 1

for sheet in xls.sheet_names:
    if sheet.lower().startswith("sheet") or sheet.lower().startswith("details"):
        continue

    df = pd.read_excel(RAW_INCIDENT_FILE, sheet_name=sheet, header=None)
    full_text = " ".join(df.astype(str).fillna("").values.flatten())

    record = {
        "Incident_ID": f"{sheet}_{incident_counter}",
        "Incident_Date": extract_date(full_text),
        "Case_Type": classify_case(full_text),
        "Warehouse": sheet.strip(),
        "Area": extract_field("Area", full_text),
        "Process": extract_field("Process", full_text),
        "Body_Part": extract_field("Body", full_text),
        "Description": full_text,
        "Source_Sheet": sheet
    }

    records.append(record)
    incident_counter += 1

incident_fact = pd.DataFrame(records)

# ----------------------------
# SAVE INCIDENT FACT
# ----------------------------
incident_fact.to_excel(
    OUT_DIR / "incident_fact.xlsx",
    index=False
)

# ----------------------------
# CREATE MAN-HOURS FACT (BASE)
# ----------------------------
manhours = (
    incident_fact[["Warehouse"]]
    .drop_duplicates()
    .assign(Man_Hours=0)
)

manhours.to_excel(
    OUT_DIR / "manhours_fact.xlsx",
    index=False
)

print("ETL completed successfully.")
print("Files generated:")
print(" - data/processed/incident_fact.xlsx")
print(" - data/processed/manhours_fact.xlsx")
# ----------------------------