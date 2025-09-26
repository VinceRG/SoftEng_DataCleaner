# reshape_to_numeric.py

import os
import pandas as pd


# CONFIG

input_file = os.path.join("cleanExcel", "cleanedBook.xlsx")
output_file = os.path.join("cleanExcel", "numericBook.xlsx")


# LOAD FILE

final_df = pd.read_excel(input_file)


# CREATE MAPPING DICTIONARIES


# Sex encoding
sex_map = {"Male": 1, "Female": 0}

# Age range encoding
age_map = {
    "Under 1": 0,
    "1-4": 1,
    "5-9": 2,
    "10-14": 3,
    "15-18": 4,
    "19-24": 5,
    "25-29": 6,
    "30-34": 7,
    "35-39": 8,
    "40-44": 9,
    "45-49": 10,
    "50-54": 11,
    "55-59": 12,
    "60-64": 13,
    "65-69": 14,
    "70": 15,
    "70 Over": 15,
    "70 & OVER": 15
}

# Consultation_Type encoding
consult_map = {name: idx for idx, name in enumerate(final_df["Consultation_Type"].dropna().unique(), start=1)}

# Case encoding
case_map = {name: idx for idx, name in enumerate(final_df["Case"].dropna().unique(), start=1)}


# RESHAPE INTO LONG FORMAT

mapping_dict = {}
for col in final_df.columns[3:]:  # skip Month_year, Consultation_Type, Case
    parts = col.split()
    if len(parts) >= 2:
        sex = parts[-1]
        age = " ".join(parts[:-1])
        mapping_dict[col] = {"Age_range": age, "Sex": sex}

reshaped_df = final_df.melt(
    id_vars=["Month_year", "Consultation_Type", "Case"],
    value_vars=final_df.columns[3:],
    var_name="Age_Sex",
    value_name="Total"
)

reshaped_df["Age_range"] = reshaped_df["Age_Sex"].map(lambda x: mapping_dict[x]["Age_range"])
reshaped_df["Sex"] = reshaped_df["Age_Sex"].map(lambda x: mapping_dict[x]["Sex"])


# SPLIT MONTH_YEAR INTO NUMERIC MONTH + YEAR

reshaped_df["Month_year"] = pd.to_datetime(reshaped_df["Month_year"], errors="coerce")
reshaped_df["Month"] = reshaped_df["Month_year"].dt.month
reshaped_df["Year"] = reshaped_df["Month_year"].dt.year


# ENCODE TO NUMERIC

reshaped_df["Sex"] = reshaped_df["Sex"].map(sex_map).fillna(-1).astype(int)
reshaped_df["Age_range"] = reshaped_df["Age_range"].map(age_map).fillna(-1).astype(int)
reshaped_df["Consultation_Type"] = reshaped_df["Consultation_Type"].map(consult_map).fillna(-1).astype(int)
reshaped_df["Case"] = reshaped_df["Case"].map(case_map).fillna(-1).astype(int)


# FINAL NUMERIC STRUCTURE

reshaped_df = reshaped_df[["Year", "Month", "Consultation_Type", "Case", "Sex", "Age_range", "Total"]]


# SAVE TO NEW EXCEL

reshaped_df.to_excel(output_file, index=False)

print(f"✅ Numeric Excel saved as: {output_file}")
print("\n🔑 Case encoding dictionary:")
print(case_map)

