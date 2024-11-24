import pandas as pd
import re

# Load the Excel file
input_file = "dataset1.xlsx"  # Replace with your Excel file name
output_file = "formatted_data1.xlsx"  # Name for the cleaned output file

# Read the Excel file without assuming a column name
df_raw = pd.read_excel(input_file, header=None)  # No header assumed

# Check if the file has data
if df_raw.empty:
    raise ValueError("The input Excel file is empty or could not be read.")

# Assuming the raw data is in the first column
raw_data_column = df_raw.iloc[:, 0]  # Take the first (and only) column

# Process the data
def extract_field(pattern, text, default=None):
    match = re.search(pattern, text, re.MULTILINE)
    return match.group(1).strip() if match else default

def process_data(raw_column):
    rows = []
    for record in raw_column:
        if not isinstance(record, str):  # Skip non-string records
            continue
        admission_date = extract_field(r"Admission Date:\s*\[\*\*(.*?)\*\*\]", record)
        discharge_date = extract_field(r"Discharge Date:\s*\[\*\*(.*?)\*\*\]", record)
        dob = extract_field(r"Date of Birth:\s*\[\*\*(.*?)\*\*\]", record)
        sex = extract_field(r"Sex:\s*(\w+)", record)
        service = extract_field(r"Service:\s*(.*?)\n", record)
        allergies = extract_field(r"Allergies:\n(.*?)\n\n", record)
        diagnosis = extract_field(r"Discharge Diagnosis:\n(.*?)\n\n", record)
        condition = extract_field(r"Discharge Condition:\n(.*?)\n\n", record)
        medications = extract_field(r"Discharge Medications:\n(.*?)(?=\n\n|\Z)", record)
        instructions = extract_field(r"Discharge Instructions:\n(.*?)(?=\n\n|\Z)", record)

        # Append extracted data as a dictionary
        rows.append({
            "Admission Date": admission_date,
            "Discharge Date": discharge_date,
            "Date of Birth": dob,
            "Sex": sex,
            "Service": service,
            "Allergies": allergies,
            "Diagnosis": diagnosis,
            "Condition": condition,
            "Medications": medications,
            "Instructions": instructions
        })
    
    return rows

# Process the raw data column
cleaned_rows = process_data(raw_data_column)

# Convert the cleaned rows to a DataFrame
df_cleaned = pd.DataFrame(cleaned_rows)

# Save the cleaned data to a new Excel file
df_cleaned.to_excel(output_file, index=False)
print(f"Data has been formatted and saved to '{output_file}'.")
