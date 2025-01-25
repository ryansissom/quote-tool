import pandas as pd
import xlwings as xw
import tkinter as tk
from tkinter import simpledialog
from fuzzywuzzy import process
from tkinter import messagebox
import os


def showDialogBox(cust_desc, options, num_options):
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    while True:  # Loop until valid input is received
        dialog = tk.Toplevel(root)
        dialog.title("Fuzzy Match Selection")

        selected_option = None

        def on_ok():
            nonlocal selected_option
            selected_option = entry.get()
            dialog.destroy()

        def on_skip():
            nonlocal selected_option
            selected_option = "skip"
            dialog.destroy()

        def on_cancel():
            nonlocal selected_option
            selected_option = "cancel"
            dialog.destroy()

        tk.Label(dialog, text=f"Customer Description: {cust_desc}").pack(pady=5)
        tk.Label(dialog, text=f"Options:\n\n{options}").pack(pady=5)

        entry = tk.Entry(dialog)
        entry.pack(pady=5)

        tk.Button(dialog, text="OK", command=on_ok).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(dialog, text="Skip", command=on_skip).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(dialog, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=5, pady=5)

        root.wait_window(dialog)  # Wait until the dialog is closed

        # Validate input
        if selected_option == "cancel" or selected_option == "skip":
            return selected_option
        elif selected_option and selected_option.isdigit() and 1 <= int(selected_option) <= num_options:
            return selected_option
        else:
            messagebox.showerror("Invalid Input", "Please enter a valid number (1–4).")

def handle_blank_cell():
    # Create a simple Tkinter dialog
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Prompt the user
    response = messagebox.askyesno(
        "Blank Cell Detected",
        "A blank cell has been detected.\n\nAre you done? (Yes to stop, No to skip)"
    )
    return response  # True for "Yes", False for "No"


def fuzzyMatch():
    wb = xw.Book('demo.xlsm')
    sheet = wb.sheets[0]
    customer_descriptions = sheet.range('A16:A2000').value
    master_df = pd.read_csv(os.path.join(os.path.dirname(__file__), "Store Parts Inventory.csv"))
    master_descriptions = master_df['Description'].tolist()

    for i, cust_desc in enumerate(customer_descriptions, start=16):
        if cust_desc is None:  # Check for a blank cell
            response = handle_blank_cell()
            if response:  # User clicked "Yes" (stop the process)
                print("Process ended by user.")
                break
            else:  # User clicked "No" (skip the line)
                print(f"Row {i} is blank. Skipping...")
                continue
        matches = process.extract(cust_desc, master_descriptions, limit=4)
        detailed_matches = []
        for match in matches:
            matched_row = master_df[master_df['Description'] == match[0]].iloc[0]
            description = matched_row['Description']
            manufacturer = matched_row['Provider']
            part_number = matched_row['Part Number']
            weighted_cost = matched_row['Weighted Average Cost']
            detailed_matches.append((description, manufacturer, part_number, weighted_cost))

        options = "\n".join([
            f"{idx + 1}. Description: {match[0]}\n Manufacturer: {match[1]}\n Part Number: {match[2]}\n"
            for idx, match in enumerate(detailed_matches)
        ])

        selected_option = showDialogBox(cust_desc, options, len(detailed_matches))
        if selected_option == "cancel":  # User pressed Cancel
            print("Quote process canceled by user.")
            break
        elif selected_option == "skip":  # User pressed Skip
            print(f"Row {i} skipped by user.")
            continue
        selected_match = detailed_matches[int(selected_option) - 1]

        # Write back to Excel: Place the values in the appropriate columns
        sheet.range(f"C{i}").value = selected_match[1]  # Manufacturer
        sheet.range(f"D{i}").value = selected_match[2]  # Part Number
        sheet.range(f"B{i}").value = i - 15  # Line Number
        sheet.range(f"E{i}").value = selected_match[0]  # Description
        sheet.range(f"H{i}").value = selected_match[3]  # Weighted Average Cost

def calculate():
    wb = xw.Book('demo.xlsm')
    sheet = wb.sheets[0]
    quantities = sheet.range('F16:F2000').value

    for i, qty in enumerate(quantities, start=16):
        cost = sheet.range(f'H{i}').value
        margin = sheet.range(f'I{i}').value

        # Check for blank cells
        if cost is None or margin is None:
            response = handle_blank_cell()
            if response:  # User clicked "Yes" (stop the process)
                print("Process ended by user.")
                break
            else:  # User clicked "No" (skip the line)
                print(f"Blank cell in row {i}. Skipping...")
                continue

        # Calculate price and total price
        price = float(cost) * (1 + float(margin))
        sheet.range(f'J{i}').value = round(price, 2)  # Price
        sheet.range(f'K{i}').value = round(price, 2) * qty  # Total price


# To do
