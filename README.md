# Distance Calculator for Google Sheets

## Files
- `distance_calculator_blr.js` → Bangalore sheet (K, L → M)
- `distance_calculator_mumbai.js` → Mumbai sheet (Q, R → S)

## How to use
1. Open Google Sheet → Extensions → Apps Script
2. Paste the relevant city file
3. Save → Refresh sheet
4. Click 🚗 Distance Tools → Fill Missing Distances

## Config (top of each file)
Change column numbers here if sheet structure changes:
COL_STORE_LAT_LONG → input column 1
COL_STARTING_POINT → input column 2  
COL_DISTANCE → output column
DATA_START_ROW → row where data starts (usually 2)
