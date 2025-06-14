#!/usr/bin/env python3

import pandas as pd
import re
from tkinter import Tk, filedialog

def select_excel_file(prompt):
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=prompt, filetypes=[("Excel files", "*.xlsx *.xls")])
    root.destroy()
    return file_path

def parse_libraries(raw_value):
    if pd.isna(raw_value):
        return []
    parts = str(raw_value).split(",")
    libraries = []
    for val in parts:
        
        val = val.replace("Reserves", "Library").strip()
        match = re.match(r"^(Ginn|HHSL|Music|SMFA|Tisch|Veterinary|Biology|Geology|Chemistry)", val)
        if match:
            library = match.group(1)
            if library in ['Biology', 'Geology', 'Chemistry']:
                library = "Tisch"
            if library == "Veterinary":
                library = "Vet"
            
            print(library)
            libraries.append(library)
    return libraries

def build_from_mapping(df_mapping):
    df_mapping['Library-Item Policy'] = df_mapping['Library Name'] + "-" + df_mapping['Current Item Policy']

    return df_mapping['Library-Item Policy'].unique().tolist()
def build_extant_set_from_pivot(df_pivot):
    extant_set = []
    df_pivot = df_pivot.fillna(0)

    for item_policy in df_pivot.index:
        for library in df_pivot.columns:
            try:
                val = float(df_pivot.at[item_policy, library])
                if val > 0:
                    extant_set.append(f"{library.strip()}-{item_policy.strip()}")
            except Exception:
                continue
    return extant_set


def main():
    loan_rules_file = select_excel_file("Select the Loan Rules Excel file")
    mapping_file = select_excel_file("Select the Item Policy Mapping Excel file")
    pivot_file = select_excel_file("Select the Pivot Table for Extant Library-Item Policy Combinations")

    df_loan_rules = pd.read_excel(loan_rules_file)
    df_mapping = pd.read_excel(mapping_file)
    df_pivot = pd.read_excel(pivot_file)
    df_pivot.columns = df_pivot.columns.map(str)
    df_pivot["Item Policy"] = df_pivot["Item Policy"].astype(str).str.strip()
    df_pivot = df_pivot.set_index("Item Policy")


    required_loan_cols = {"Item Policy Value", "Location Value", "Possible Locations"}
    required_mapping_cols = {"Library Name", "Current Item Policy", "New item policy/Loan Length"}

    if not required_loan_cols.issubset(df_loan_rules.columns):
        raise ValueError(f"Loan rules file must contain columns: {required_loan_cols}")
    if not required_mapping_cols.issubset(df_mapping.columns):
        raise ValueError(f"Mapping file must contain columns: {required_mapping_cols}")

    # Build lookup dicts
    mapping_dict = {
        str(row["Library Name"]).strip() + "-" + str(row["Current Item Policy"]).strip():
        str(row["New item policy/Loan Length"]).strip()
        for _, row in df_mapping.iterrows()
    }

    extant_set = build_extant_set_from_pivot(df_pivot)
    already_mapped = build_from_mapping(df_mapping)
    non_extant = []
    unmapped = []

    def map_new_item_policies(row):
        # Get item policy values
        item_policies = str(row["Item Policy Value"]).split(",") if pd.notna(row["Item Policy Value"]) else []

        # Get list of parsed libraries from "Location Value", or fallback to "Possible Locations"
        libraries = parse_libraries(row["Location Value"])
        if not libraries:
            libraries = parse_libraries(row["Possible Locations"])
        if not libraries or not item_policies:
            return []

        new_policies = []
        for policy in map(str.strip, item_policies):
            print(policy)
            for lib in libraries:
                
                combo_key = f"{lib.strip()}-{policy.strip()}"

                if combo_key in extant_set:
                    if combo_key in already_mapped:
                        print(mapping_dict)
                        new_item_policy = mapping_dict[combo_key]
                        print(new_item_policy)

                        new_policies.append(new_item_policy)
                        
                    else:
                        new_policies.append(f"(Unmapped: {lib}-{policy})")
                        unmapped.append(f"(Unmapped: {lib}-{policy})")
                        
                else:
                    continue
                    
        return new_policies

    # Apply to dataframe
    df_loan_rules["New Item Policy"] = df_loan_rules.apply(map_new_item_policies, axis=1)

    # Save output
    output_path = loan_rules_file.replace(".xlsx", "_with_new_item_policy.xlsx")
    df_loan_rules.to_excel(output_path, index=False)

    with open("Non-Extant Mappings.txt", "w") as f:
        f.write("\n".join(sorted(set(non_extant))))

    with open("Unmapped Mappings.txt", "w") as f:
        f.write("\n".join(sorted(set(unmapped))))

    print(f"✅ Saved: {output_path}")
    print(f"⚠️  Non-extant combos saved to: Non-Extant Mappings.txt")
    print(f"❓ Unmapped combos saved to: Unmapped Mappings.txt")

if __name__ == "__main__":
    main()
