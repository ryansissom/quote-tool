# Quote Tool

## Description
The Quote Tool automates the process of creating quotes in Excel. It matches customer descriptions to inventory items, calculates prices, and manages margins with minimal manual effort.

---

## Features
- **Fuzzy Matching**: Matches customer descriptions to the closest inventory items.
- **Interactive Selection**: Dialogs allow users to select, skip, or cancel matches.
- **Automatic Price Calculation**: Computes prices and total costs using provided quantities and margins.

---

## Prerequisites
1. **Python Installation**:
   - Ensure Python is installed.
   - Install dependencies using:
     ```bash
     pip install pandas xlwings tkinter fuzzywuzzy python-Levenshtein
     ```
2. **Required Files**:
   - Place `demo.xlsm` and `Store Parts Inventory.csv` in the same directory as the script.

---

## Usage

### Prepare the Excel File:
1. Open `demo.xlsm`.
2. Input customer descriptions in **column A** (starting from row 16).

### Run the Script:
1. **Run `fuzzyMatch()`**:
   - Matches descriptions and populates columns:
     - **B**: Line Number
     - **C**: Manufacturer
     - **D**: Part Number
     - **E**: Description
     - **H**: Weighted Average Cost
2. **Input Quantities and Margins**:
   - Fill **column F** with quantities.
   - Fill **column I** with desired margins.
3. **Run `calculate()`**:
   - Fills columns:
     - **J**: Price
     - **K**: Total Price

---

## Example Workflow
1. Enter descriptions in **column A**.
2. Run `fuzzyMatch()` to populate necessary columns.
3. Input quantities and margins in columns **F** and **I**.
4. Run `calculate()` to generate prices and total costs.

---

## Notes
- If a blank cell is detected, the script will prompt you to skip or stop.
- Canceling `fuzzyMatch()` stops the process entirely.
- **For Distribution to Non-Technical Users**:
  - Use [PyInstaller](https://pyinstaller.org) to package the script as a `.exe` (Windows).
  - On macOS, use the xlwings add-in for seamless integration.

---

## Support
For troubleshooting or questions, contact **[Your Name/Email]**.
