# ui_functions.py

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
    
    # --- THIS IS THE LOGIC THAT WAS MISSING ---
    has_name_col_raw = input("Does your BOQ have separate 'Name' and 'Description' columns? (yes/no): ").lower()
    has_name_col = True if has_name_col_raw == 'yes' else False

    mapping = {}
    
    if has_name_col:
        standard_fields = ['Name', 'Description', 'Quantity', 'Rate', 'UOM']
    else:
        standard_fields = ['Description', 'Quantity', 'Rate', 'UOM']
    # ---------------------------------------------

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