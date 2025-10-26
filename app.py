import os
import pandas as pd
import openpyxl # is used by pandas, but it is required as a separate library to set column widths
from pathlib import Path


##### Functions
def move_files(folder_path, destination_folder):
    # Initialize variables
    filesMoved = 0

    ##### Move "Otros" csv files to new destination folder and convert them to xlsx
    for file in Path(folder_path).rglob('*.csv'):  # Use rglob to search recursively in subfolders
        # Read the CSV file with the correct encoding
        df = pd.read_csv(file, encoding='ISO-8859-1')
        
        # Get the relative path from the source folder
        relative_path = file.relative_to(folder_path)

        # Get first part of the relative path
        first_part_relative_path = relative_path.parts[0]

        if first_part_relative_path == "Compensaciones":
            # Add "Facturacion y Cobranza" at the start of the relative path
            relative_path = Path("Facturacion y Cobranza") / relative_path

        # Create the destination path preserving the subfolder structure
        dest_file_path = Path(destination_folder) / relative_path.with_suffix('.xlsx')
        
        # Create the destination directory if it doesn't exist
        dest_file_path.parent.mkdir(parents=True, exist_ok=True)

        # Check if the name of the file start with a value from this list to see if there are changes that need to be made. This is the list: ["FacturasPER", "FacturasPAN"] 
        if file.name.startswith("FacturasPAN"):
            # Change the values in the column "Mon" to "USD"
            df["Mon"] = "USD"
        elif file.name.startswith("FacturasPER"):
            # Change the values in the column "Mon" to "SOL"
            df["Mon"] = "SOL"
        elif file.name.startswith("CobranzaPAN"):
            # Change the values in the column "FDE_MONEDA" and "FAC_MONEDA" to "USD"
            df["FDE_MONEDA"] = "USD"
            df["FAC_MONEDA"] = "USD"
        elif file.name.startswith("CobranzaPER"):
            # Change the values in the column "FDE_MONEDA" and "FAC_MONEDA" to "SOL"
            df["FDE_MONEDA"] = "SOL"
            df["FAC_MONEDA"] = "SOL"
        elif file.name.startswith("CompensacionesPAN"):
            # Change the values in the column "Moneda" to "USD"
            df["Moneda"] = "USD"
        elif file.name.startswith("CompensacionesPER"):
            # Change the values in the column "Moneda" to "SOL"
            df["Moneda"] = "SOL"
        elif file.name.startswith("NotasCrePAN"):
            # Change the values in the column "NCRE_MONEDA" to "USD"
            df["NCRE_MONEDA"] = "USD"
        elif file.name.startswith("NotasCrePER"):
            # Change the values in the column "NCRE_MONEDA" to "SOL"
            df["NCRE_MONEDA"] = "SOL"
        
        # Save the Excel file
        df.to_excel(dest_file_path, index=False)
        
        # Remove the original CSV file
        file.unlink()  # Delete the original CSV file

        filesMoved += 1

    return filesMoved

def convert_pending_csv_to_xlsx(destination_folder):
    # Initialize variables
    filesConverted = 0

    ##### Convert pending csv files to xlsx
    for file in Path(destination_folder).rglob('*.csv'):
        print(file)

        with open(file, encoding='ISO-8859-1') as f:
            for i, line in enumerate(f):
                if i in range(265, 275):
                    print(i, line)


        # Read the CSV file with the correct encoding
        df = pd.read_csv(file, encoding='ISO-8859-1', sep=',', quotechar='"', engine='python')
        
        # Save the Excel file
        df.to_excel(file.with_suffix('.xlsx'), index=False)

        # Remove the original CSV file
        file.unlink()  # Delete the original CSV file

        filesConverted += 1

    return filesConverted


##### Main script
# Initialize variables
files_folder = "C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/Otros" # C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/Otros     # C:/Users/pc/Desktop/ba-files/Otros
destination_folder = "C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/DYCUSA" # C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/DYCUSA     # C:/Users/pc/Desktop/ba-files/DYCUSA

# Move "Otros" files to new folder
filesMoved = move_files(files_folder, destination_folder)
if filesMoved > 0:
    print(f"Files moved successfully: {filesMoved}")
else:
    print("No files where moved")

# Convert pending csv files to xlsx
extraFilesConverted = convert_pending_csv_to_xlsx(destination_folder)
if extraFilesConverted > 0:
    print(f"Extra files converted successfully: {extraFilesConverted}")
else:
    print("No extra files where converted")