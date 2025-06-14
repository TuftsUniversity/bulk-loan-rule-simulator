#!/usr/bin/env python3

import pandas as pd
from tkinter import Tk, filedialog

def select_file(prompt, filetypes):
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=prompt, filetypes=filetypes)
    root.destroy()
    return file_path

def main():
    # Prompt for the mapping file
    mapping_path = select_file("Select the Item Policy Mapping Excel file", [("Excel files", "*.xlsx *.xls")])
    if not mapping_path:
        print("Mapping file not selected. Exiting.")
        return

    # Prompt for the unmapped values text file
    unmapped_path = select_file("Select the Unmapped Mappings TXT file", [("Text files", "*.txt")])
    if not unmapped_path:
        print("Unmapped mappings file not selected. Exiting.")
        return

    # Load existing mapping file
    df_mapping = pd.read_excel(mapping_path)

    # Read unmapped entries
    with open(unmapped_path, "r") as f:
        lines = [line.strip() for line in f if line.strip() and "-" in line]

    # Build new rows
    new_rows = []
    for line in lines:
        library_name, policy = line.split("-", 1)
        new_rows.append({
            "Library Name": library_name,
            "Library Code": library_name.upper(),
            "Current Item Policy": policy,
            "Item Policy Code": policy.replace(" ", ""),
            "New item policy/Loan Length": "",  # leave others blank
        })

    # Create DataFrame and ensure all columns exist
    df_new = pd.DataFrame(new_rows)
    for col in df_mapping.columns:
        if col not in df_new.columns:
            df_new[col] = ""

    # Reorder to match original
    df_new = df_new[df_mapping.columns]

    # Concatenate and save
    df_result = pd.concat([df_mapping, df_new], ignore_index=True)
    output_path = mapping_path.replace(".xlsx", "_with_Unmapped.xlsx")
    df_result.to_excel(output_path, index=False)

    print(f"✅ New rows added and saved as: {output_path}")

if __name__ == "__main__":
    main()
