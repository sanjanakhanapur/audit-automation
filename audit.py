import pandas as pd
import numpy as np
import logging
import sys

# ---------------- CONFIG ----------------
AUDIT_RULES = {
    "Contact Owner Present": ("Contact owner", "notna"),
    "Owner Assigned Date Present": ("Owner assigned date", "notna"),
    "BDR SLA Met": ("SLA ALERT BDR", "isna"),
    "AM SLA Met": ("SLA ALERT AM", "isna"),
    "Lead Source Present": ("Lead Source", "notna"),
}

PASS_SCORE = 80

# ------------- LOGGING ------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ------------- INPUT FILES --------------
input_file = sys.argv[1] if len(sys.argv) > 1 else "hubspot_export.xlsx"
output_file = sys.argv[2] if len(sys.argv) > 2 else "audit_result.xlsx"

logging.info(f"Reading input file: {input_file}")

# ------------- READ DATA ----------------
df = pd.read_excel(input_file)

# Keep original column order
original_columns = df.columns.tolist()

# ------------- AUDIT CHECKS -------------
for audit_col, (source_col, rule) in AUDIT_RULES.items():
    if rule == "notna":
        df[audit_col] = np.where(df[source_col].notna(), "Yes", "No")
    elif rule == "isna":
        df[audit_col] = np.where(df[source_col].isna(), "Yes", "No")

audit_columns = list(AUDIT_RULES.keys())

# ------------- FAIL COUNT ----------------
df["Fail Count"] = (df[audit_columns] == "No").sum(axis=1)

# ------------- SCORE ---------------------
TOTAL_CHECKS = len(AUDIT_RULES)
POINTS_PER_FAIL = 100 // TOTAL_CHECKS

df["Audit Score"] = 100 - (df["Fail Count"] * POINTS_PER_FAIL)
df["Audit Score"] = df["Audit Score"].clip(lower=0)

# ------------- PASS / FAIL ---------------
df["Audit Result"] = np.where(df["Audit Score"] >= PASS_SCORE, "PASS", "FAIL")

# ------------- FAILURE REASON ------------
def failure_reason(row):
    failed_checks = [col for col in audit_columns if row[col] == "No"]
    return ", ".join(failed_checks) if failed_checks else "None"

df["Failure Reason"] = df.apply(failure_reason, axis=1)

# ------------- COLUMN ORDER --------------
final_columns = (
    original_columns +
    audit_columns +
    ["Fail Count", "Audit Score", "Audit Result", "Failure Reason"]
)

df = df[final_columns]

# ------------- SAVE OUTPUT ---------------
df.to_excel(output_file, index=False)

logging.info(f"Audit completed successfully. Output saved as {output_file}")