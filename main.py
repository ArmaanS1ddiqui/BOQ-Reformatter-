# main.py
import pandas as pd
import os
from ui_functions import get_sheet, find_header, get_cols, map_cols
from processing_functions import clean_df, create_sections, trim_df

# main.py

def main():
    """Main function to orchestrate the BOQ cleaning workflow."""
    
    # --- HARDCODE YOUR FILE PATH HERE ---
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
    
    #inspecting the data before creating sections
    print("\n--- Data being sent to create_sections (DEBUG PREVIEW) ---")
    print(cleaned_data.head(25)) # Print the first 25 rows
    
    data_with_sections = create_sections(cleaned_data, field_mapping)
    final_boq = trim_df(data_with_sections, field_mapping)
    
    print("\n✅ --- Processing Complete! ---")
    print("Here is a preview of your final cleaned BOQ data:")
    print(final_boq.head())
    
    save_choice = input("\nDo you want to save this to a CSV file? (yes/no): ").lower()
    if save_choice == 'yes':
        output_folder = "Output"
        os.makedirs(output_folder, exist_ok=True)
        
        base_filename = os.path.basename(filepath)
        new_filename = os.path.splitext(base_filename)[0] + "_cleaned.csv"
        
        output_path = os.path.join(output_folder, new_filename)
        
        final_boq.to_csv(output_path, index=False)
        print(f"File saved successfully to: {output_path}")

if __name__ == "__main__":
    main()