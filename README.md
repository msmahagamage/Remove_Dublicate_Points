# Excel Deduplication with Spatial Clustering + Fuzzy Address Matching

This project combines multiple Excel files from a folder (including subfolders) into one dataset and removes likely duplicate records using a **hybrid deduplication approach**:

1. **Spatial clustering**: groups records within a **200-meter** radius.
2. **Hybrid matching inside each cluster**:
   - **Exact match** on **street number**
   - **Fuzzy match** on **street name** and **store name** (string similarity thresholds)

It produces:
- A consolidated Excel file with **unique** records
- A comparison Excel file showing which rows were treated as **original vs removed duplicates**

---

## Features

- Recursively reads `.xlsx` and `.xls` files from a root directory
- Standardizes inconsistent column names across files
- Cleans coordinates and text fields
- Extracts street number and street name from addresses
- Detects duplicates using spatial proximity + fuzzy matching
- Saves two outputs:
  - `Consolidated_Unique_Data11.xlsx`
  - `Removed_Duplicates_Comparison11.xlsx`

---

## Requirements

### Python
- Python 3.9+ recommended

### Libraries
Install required packages:

```bash
pip install pandas openpyxl geopy thefuzz python-Levenshtein numpy
