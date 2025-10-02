import pandas as pd
import re

# ----------------------------
# Normalization helpers
# ----------------------------
# Function to clean and standardize SKU values
def normalize_sku_value(x: object) -> str | None:
    if pd.isna(x): # If the value is NaN, return None
        return None
    s = str(x).replace("\u00A0", " ").strip() # Convert to string, replace non-breaking spaces, trim
    if s == "" or s.upper() in {"NAN", "NONE", "NULL"}: # Handle empty or placeholder values
        return None
    # If the value looks like a number (e.g. '165080.0'), convert to integer string
    if re.fullmatch(r"\d+(\.0+)?", s):
        try:
            n = int(float(s)) # Convert to float then int
            return str(n) # Return as string
        except Exception:
            pass
    # Otherwise, collapse internal whitespace and uppercase
    s = re.sub(r"\s+", " ", s).upper()
    return s

# Applies normalization to an entire pandas Series
def normalize_series(series: pd.Series) -> pd.Series:
    return series.map(normalize_sku_value, na_action="ignore")

# ----------------------------
# Load File A (source dictionary)
# ----------------------------
file_a_path = r"C:\Users\lgabriel\Downloads\ACOMPANHAMENTO FCST PERFORMANCE - Setembro.xlsx" # Path to the source Excel file

# Reads columns E and R from File A, treating column E as string
file_a = pd.read_excel(
    file_a_path,
    usecols="E,R", #first column is for the sku second column is for the value that will be picked
    dtype={"E": "string"},
    engine="openpyxl"
)
file_a.columns = ["SKU_RAW", "VALUE_R"] # Renames columns for clarity

file_a["SKU_KEY"] = normalize_series(file_a["SKU_RAW"]) # Normalizes SKUs
# Removes rows with missing SKUs and drops duplicates, keeping the first occurrence
file_a = file_a.dropna(subset=["SKU_KEY"]).drop_duplicates(subset=["SKU_KEY"], keep="first")

# Creates a dictionary mapping normalized SKUs to their values
sku_dict = dict(zip(file_a["SKU_KEY"], file_a["VALUE_R"]))

# ----------------------------
# Load File B (multiple sheets)
# ----------------------------
import openpyxl # Imports openpyxl for direct Excel editing

file_b_path = r"C:\Users\lgabriel\Downloads\TodasAsClassesTemplate SMAPE & TS Calculation_17_07_2025.xlsx" # Path to the target workbook
wb = openpyxl.load_workbook(file_b_path) # Loads the workbook for writing
sheet_names = [ # List of sheet names to process
    "Suínos",
    "AVES",
    "RUM",
    "pet-bio (goldschmidt)",
    "pet-para (pcastro)",
    "pet-para (boliveira)",
    "Equinos"
]

file_b = pd.read_excel( # Reads column A from all sheets using pandas (for initial analysis)
    file_b_path,
    sheet_name=sheet_names,
    usecols="A",
    dtype={"A": "string"},
    engine="openpyxl"
)

block_size = 16 # Defines the size of each SKU block

# ----------------------------
# Iterate and match per block (use FIRST non-empty)
# ----------------------------
# First loop: prints sheet info and checks for data presence
for sheet_name in sheet_names:
    print(f"\n--- Sheet: {sheet_name} ---")
    df = file_b[sheet_name].copy() # Copies the sheet's data
    df.columns = ["SKU_A_RAW"] # Renames column A for clarity

    n_rows = len(df) # Gets number of rows
    if n_rows <= 1: # Skips empty sheets
        print("No data rows found.")
        continue

# Second loop: processes each sheet for writing
for sheet_name in sheet_names:
    print(f"\n📄 Sheet: {sheet_name}")
    sheet = wb[sheet_name] # Accesses the sheet in the workbook
    # Extracts column A values into a DataFrame
    df = pd.DataFrame([cell.value for cell in sheet['A']], columns=["SKU_A_RAW"])
    n_rows = len(df)

    # Loops through each block of 16 rows
    for start in range(1, n_rows, block_size):
        end = min(start + block_size, n_rows)
        block = df.iloc[start:end].copy() # Extracts the block

        # Filters out empty or whitespace-only SKUs
        non_empty = block["SKU_A_RAW"].dropna()
        non_empty = non_empty[non_empty.astype(str).str.strip() != ""]

        if non_empty.empty: # Skips blocks with no valid SKUs
            continue

        # Normalizes SKUs and keeps valid ones
        candidates = [(idx, val, normalize_sku_value(val)) for idx, val in non_empty.items()]
        candidates = [(idx, val, key) for idx, val, key in candidates if key is not None]

        if not candidates: # Skips blocks with no usable SKUs
            continue

         # Picks the first valid SKU in the block
        first_idx, first_raw, first_key = candidates[0]
        target_row = first_idx + 1  # Calculates target row (1 rows below SKU position)
        target_col = 11  # Column K

        # If SKU matches, write value to column K
        if first_key in sku_dict:
            value = sku_dict[first_key]
            sheet.cell(row=target_row, column=target_col, value=value)
            print(f"Row {target_row}: ✅ Wrote value '{value}' for SKU '{first_key}'")
        else:
            print(f"Row {target_row}: ❌ No match for SKU '{first_key}'")

from datetime import datetime  # Import datetime for timestamping

# Generate timestamp in format YYYYMMDD_HHMM
timestamp = datetime.now().strftime("%Y%m%d_%H%M")

# Create output path with timestamp
output_path = fr"C:\Users\lgabriel\Downloads\todas_as_classes_smape_{timestamp}.xlsx"

# Save the updated workbook
wb.save(output_path)

print(f"\n✅ File saved successfully as: {output_path}")