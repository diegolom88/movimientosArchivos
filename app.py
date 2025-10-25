import os
import pandas as pd
import openpyxl # is used by pandas, but it is required as a separate library to set column widths
from pathlib import Path

##### Functions
def move_files(folder_path, destination_folder):
    # initialize variables
    filesMoved = 0

    # convert csv files to xlsx files
    for file in Path(folder_path).rglob('*.csv'):  # Use rglob to search recursively in subfolders
        df = pd.read_csv(file)
        
        # Get the relative path from the source folder
        relative_path = file.relative_to(folder_path)

        # get first part of the relative path
        first_part_relative_path = relative_path.parts[0]

        if first_part_relative_path == "Compensaciones":
            # add "Facturacion y Cobranza" at the start of the relative path
            relative_path = Path("Facturacion y Cobranza") / relative_path

        # Create the destination path preserving the subfolder structure
        dest_file_path = Path(destination_folder) / relative_path.with_suffix('.xlsx')
        
        # Create the destination directory if it doesn't exist
        dest_file_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Save the Excel file
        df.to_excel(dest_file_path, index=False)
        
        # Remove the original CSV file
        file.unlink()  # Delete the original CSV file

        filesMoved += 1

    return filesMoved



##### Main script
# initialize variables
files_folder = "C:/Users/pc/Desktop/ba-files/Otros" # C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/Otros
destination_folder = "C:/Users/pc/Desktop/ba-files/DYCUSA" # C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/DYCUSA

filesMoved = move_files(files_folder, destination_folder)

if filesMoved > 0:
    print(f"Files moved successfully: {filesMoved}")
else:
    print("No files where moved")




# combined_df = combine_excel_files(folder_path)
# result = add_maquinaria_categorization(combined_df)

# # Create output folder if it doesn't exist
# os.makedirs(output_folder, exist_ok=True)

# # Save combined data to a new Excel file in the specified output folder
# output_path = os.path.join(output_folder, "MayorAcumDYCUSA.xlsx")
# print("SAVING COMBINED DATA TO EXCEL")
# result.to_excel(output_path, index=False, sheet_name="MayorAcumDYCUSA", engine="openpyxl")

# # Auto-adjust column widths
# from openpyxl import load_workbook
# wb = load_workbook(output_path)
# ws = wb["MayorAcumDYCUSA"]

# # Auto-adjust column widths based on content
# for column in ws.columns:
#     max_length = 0
#     column_letter = column[0].column_letter
#     for cell in column:
#         try:
#             if len(str(cell.value)) > max_length:
#                 max_length = len(str(cell.value))
#         except:
#             pass
#     adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
#     ws.column_dimensions[column_letter].width = adjusted_width
# wb.save(output_path)
# wb.close()
# print("FINISHED SAVING COMBINED DATA TO EXCEL")
# print(f"Files combined successfully! Saved to: {output_path}")
