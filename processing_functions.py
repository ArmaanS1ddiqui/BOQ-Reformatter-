
import pandas as pd
import re

def clean_df(df, header_idx, cols_to_keep, mapping):
    """Performs initial cleaning but keeps heading rows for the section creation step."""
    print("\n--- Step 5: Cleaning Data ---")
    data = df.iloc[header_idx:].copy()
    data.columns = data.iloc[0]
    data = data[1:]
    data = data[cols_to_keep]
    
    desc_col = mapping['Description']
    
    # Drop rows that don't have any description at all
    data.dropna(subset=[desc_col], inplace=True)

    # Remove rows with keywords like 'Sub total' or 'Note'
    keywords_to_remove = ['sub total', 'note']
    pattern = '|'.join(keywords_to_remove)
    data = data[~data[desc_col].str.lower().str.contains(pattern, na=False)].copy()
    
    
    data[desc_col] = data[desc_col].astype(str).str.strip()
    if 'Name' in mapping:
        name_col = mapping['Name']
        data[name_col] = data[name_col].astype(str).str.strip()

    
    for field in ['Quantity', 'Rate']:
        if field in mapping:
            col_name = mapping[field]
            data[col_name] = pd.to_numeric(data[col_name], errors='coerce').fillna(0)
            
    print("Initial data cleaning complete.")
    return data.reset_index(drop=True)

def create_sections(df, mapping):
    """Auto-detects parent/child headings and creates a new 'Section' column."""
    print("\n--- Step 6: Creating Sections ---")
    desc_col = mapping['Description']
    qty_col = mapping['Quantity']
    rate_col = mapping['Rate']
    
    
    parent_pattern = re.compile(r'^[A-Z]\.|^([IVXLC]+\.)') 
    child_pattern = re.compile(r'^\d+\.\d+|^[a-z]\)') 
    ambiguous_pattern = re.compile(r'^\d+\.') 

    current_parent = ""
    current_child = ""
    sections = []

    for _, row in df.iterrows():
        desc = row[desc_col]
        
        # Ensure we are checking numbers against numbers (0.0 vs 0)
        qty = float(row[qty_col])
        rate = float(row[rate_col])

        
        if qty == 0.0 and rate == 0.0 and desc:
            desc_stripped = desc.strip()
            if parent_pattern.match(desc_stripped):
                current_parent = desc_stripped
                current_child = ""
            elif child_pattern.match(desc_stripped):
                current_child = desc_stripped
            elif ambiguous_pattern.match(desc_stripped):
                if not current_parent:
                    current_parent = desc_stripped
                    current_child = ""
                else:
                    current_child = desc_stripped
        
        # Construct the section text for the current row
        if current_parent and current_child:
            sections.append(f"{current_parent} : {current_child}")
        else:
            sections.append(current_parent)

    df['Section'] = sections
    
    # remove the original heading rows
    final_df = df[(df[qty_col] != 0.0) | (df[rate_col] != 0.0)].copy()
    
    print("Sections created and heading rows removed.")
    return final_df

def trim_df(df, mapping):
    """Trims the DataFrame based on the last row with a UOM value."""
    print("\n--- Step 7: Finding BOQ End ---")
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