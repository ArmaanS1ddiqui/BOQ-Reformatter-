import pandas as pd
import os

def get_sheet(workbook):
    """Asks user for the sheet name to process."""
    sheet_names = workbook.sheet_names
    print("\n--- Step 1: Select a Sheet ---")
    print("Available sheets:", sheet_names)

    chosen_sheet = ""
    while chosen_sheet not in sheet_names:
        chosen_sheet = input("Which sheet do you want to process? ")
        if chosen_sheet not in sheet_names:
            print(f"Error: '{chosen_sheet}' is not a valid sheet name.")
    
    return chosen_sheet

def find_header(df):
    """Scans the DataFrame to locate the BOQ header row with user confirmation."""
    print("\n--- Step 2: Locating Header Row ---")
    keywords = ['description', 'particulars', 'qty', 'quantity', 'rate', 'amount', 'unit']
    
    for i, row in df.head(20).iterrows():
        row_values = [str(v).lower() for v in row.dropna()]
        matches = sum(any(kw in val for kw in keywords) for val in row_values)
        
        if matches >= 3:
            print(f"\nPotential header found on row {i+1}:")
            print(list(row.dropna()))
            confirm = input("Is this the first row of the BOQ? (yes/no): ").lower()
            if confirm == 'yes':
                print("Header confirmed.")
                return i
    
    print("Could not automatically find a header row.")
    return None

def get_cols(df, header_idx):
    """Lets the user select which columns to use."""
    print("\n--- Step 3: Select Columns ---")
    headers = df.iloc[header_idx].dropna().tolist()
    
    for i, col in enumerate(headers):
        print(f"  {i+1}: {col}")
        
    while True:
        try:
            selection = input("Which columns should be used (e.g., 2,3,4,6)? ")
            selected_indices = [int(s.strip()) - 1 for s in selection.split(',')]
            selected_columns = [headers[i] for i in selected_indices]
            print("You selected:", selected_columns)
            return selected_columns
        except (ValueError, IndexError):
            print("Invalid input. Please enter numbers from the list.")

def map_cols(columns_to_use):
    """Maps user-selected columns to standard BOQ fields, handling an optional 'Name' column."""
    print("\n--- Step 4: Map Fields ---")
    
    # --- NEW: Ask the user if a separate 'Name' column exists ---
    has_name_col_raw = input("Does your BOQ have separate 'Name' and 'Description' columns? (yes/no): ").lower()
    has_name_col = True if has_name_col_raw == 'yes' else False

    mapping = {}
    
    # --- NEW: The list of fields to map is now dynamic ---
    if has_name_col:
        standard_fields = ['Name', 'Description', 'Quantity', 'Rate','UOM']
    else:
        standard_fields = ['Description', 'Quantity', 'Rate','UOM']

    for field in standard_fields:
        print(f"\nWhich column represents '{field}'?")
        for i, col in enumerate(columns_to_use):
            print(f"  {i+1}: {col}")
        
        while True:
            try:
                choice = int(input(f"Select number for '{field}': ")) - 1
                if 0 <= choice < len(columns_to_use):
                    mapping[field] = columns_to_use[choice]
                    break
                else:
                    print("Invalid number.")
            except ValueError:
                print("Please enter a valid number.")

    print("\nFinal Mapping:", mapping)
    return mapping

def clean_df(df, header_idx, cols_to_keep, mapping):
    """Cleans the data by removing unwanted rows and formatting types."""
    print("\n--- Step 5: Cleaning Data ---")
    data = df.iloc[header_idx:].copy()
    data.columns = data.iloc[0]
    data = data[1:]
    data = data[cols_to_keep]
    
    desc_col = mapping['Description']
    qty_col = mapping['Quantity']
    rate_col = mapping['Rate']

    data.dropna(subset=[desc_col], inplace=True)

    keywords_to_remove = ['sub total', 'note']
    pattern = '|'.join(keywords_to_remove)
    data = data[~data[desc_col].str.lower().str.contains(pattern, na=False)].copy()

    data = data[(data[qty_col] != 0) | (data[rate_col] != 0)].copy()
    
    data[desc_col] = data[desc_col].astype(str).str.strip()

    # --- NEW: Clean the 'Name' column if it exists in the mapping ---
    if 'Name' in mapping:
        name_col = mapping['Name']
        data[name_col] = data[name_col].astype(str).str.strip()

    for field in ['Quantity', 'Rate', 'Amount']:
        if field in mapping:
            col_name = mapping[field]
            data[col_name] = pd.to_numeric(data[col_name], errors='coerce').fillna(0)
    
    print("Data cleaned successfully.")
    return data.reset_index(drop=True)

def trim_df(df, mapping):
    """Trims the DataFrame based on the last row with a UOM value."""
    print("\n--- Step 6: Finding BOQ End ---")
    
    if 'UOM' not in mapping:
        print("Warning: UOM column not mapped. Cannot determine a specific end point.")
        return df

    uom_col = mapping['UOM']
    non_empty_uom = df[df[uom_col].notna()]
    
    if not non_empty_uom.empty:
        last_valid_index = non_empty_uom.index[-1]
        trimmed_df = df.loc[:last_valid_index]
        print(f"BOQ trimmed. The last item with a UOM is on row {last_valid_index + 1}.")
        return trimmed_df
    
    print("Warning: Could not find any rows with a UOM value.")
    return df

def main():
    """Main function to orchestrate the BOQ cleaning workflow."""
    
    filepath = r"C:\Users\rando\Desktop\codes\Projects\BOQ_Reformatter\BOQ_Data_UAE\GreenCurve_BOQ.xlsx" 
    
    print("Starting the Interactive BOQ Cleaner.")
    print(f"Processing file: {filepath}")

    if not os.path.exists(filepath):
        print("Error: File not found. Please check the hardcoded 'filepath'.")
        return

    try:
        workbook = pd.ExcelFile(filepath)
        df_dict = {sheet: workbook.parse(sheet, header=None) for sheet in workbook.sheet_names}
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    sheet_name = get_sheet(workbook)
    df = df_dict[sheet_name]
    
    header_index = find_header(df)
    if header_index is None: return
        
    cols_to_use = get_cols(df, header_index)
    field_mapping = map_cols(cols_to_use)
    
    cleaned_data = clean_df(df, header_index, cols_to_use, field_mapping)
    final_boq = trim_df(cleaned_data, field_mapping)
    
    print("\n✅ --- Cleaning Complete! ---")
    print("Here is a preview of your final cleaned BOQ data:")
    print(final_boq.head())
    
    save_choice = input("\nDo you want to save this to a CSV file? (yes/no): ").lower()
    if save_choice == 'yes':
        # --- NEW: Logic to save the file in an 'output' folder ---
        output_folder = "output"
        os.makedirs(output_folder, exist_ok=True) # Create folder if it doesn't exist
        
        base_filename = os.path.basename(filepath) # e.g., "GreenCurve_BOQ.xlsx"
        new_filename = os.path.splitext(base_filename)[0] + "_cleaned.csv" # e.g., "GreenCurve_BOQ_cleaned.csv"
        
        output_path = os.path.join(output_folder, new_filename)
        
        final_boq.to_csv(output_path, index=False)
        print(f"File saved successfully to: {output_path}")

if __name__ == "__main__":
    main()