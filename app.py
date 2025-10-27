import os
import pandas as pd
import openpyxl # is used by pandas, but it is required as a separate library to set column widths
from pathlib import Path
from datetime import datetime


##### Functions
def update_exchange_rate_file(date_column, currency_column, file_df):
    # Initialize variables
    added_combinations = 0
    exchange_rate_file_path = "C:/Users/pc/Desktop/ba-files/DYCUSA/Datos/TiposDeCambioOtrasExtensiones.xlsx"  # C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/DYCUSA/Datos/TiposDeCambioOtrasExtensiones.xlsx     # C:/Users/pc/Desktop/ba-files/DYCUSA/Datos/TiposDeCambioOtrasExtensiones.xlsx
    
    # Open the Excel file
    currency_exchange_rate_df = pd.read_excel(exchange_rate_file_path)

    # Change name of the date_column to "Fecha" and the currency_column to "Moneda"
    file_df = file_df.rename(columns={date_column: "Fecha", currency_column: "Moneda"})

    # Identify missing (Fecha, Moneda) combinations
    merged = file_df.merge(
        currency_exchange_rate_df[["Fecha", "Moneda"]],
        on=["Fecha", "Moneda"],
        how="left",
        indicator=True,
    )

    missing_combinations = merged[merged["_merge"] == "left_only"][["Fecha", "Moneda"]]

    # remove duplicates from missing_combinations
    missing_combinations = missing_combinations.drop_duplicates()

    # If there are missing combinations, add them to the currency_exchange_rate_df
    if not missing_combinations.empty:
        # Add the missing combinations to the currency_exchange_rate_df
        currency_exchange_rate_df = pd.concat([currency_exchange_rate_df, missing_combinations])
        # Save the currency exchange rate dataframe to the excel file
        currency_exchange_rate_df.to_excel(exchange_rate_file_path, index=False)
        # Update the number of added combinations
        added_combinations = len(missing_combinations)

    # Print the number of added combinations
    print(f"Added {added_combinations} combinations to the exchange rate file")

    # Return the number of added combinations
    return added_combinations

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
            # Add missing combinations to the exchange rate file
            added_combinations = update_exchange_rate_file(date_column="Fecha", currency_column="Mon", file_df=df[["Fecha", "Mon"]])
        elif file.name.startswith("FacturasPER"):
            # Change the values in the column "Mon" to "SOL"
            df["Mon"] = "SOL"
            # Add missing combinations to the exchange rate file
            added_combinations = update_exchange_rate_file(date_column="Fecha", currency_column="Mon", file_df=df[["Fecha", "Mon"]])
        elif file.name.startswith("CobranzaPAN"):
            # Change the values in the column "FDE_MONEDA" and "FAC_MONEDA" to "USD"
            df["FDE_MONEDA"] = "USD"
            df["FAC_MONEDA"] = "USD"
            # Add missing combinations to the exchange rate file
            added_combinations = update_exchange_rate_file(date_column="FDE_FECHA", currency_column="FDE_MONEDA", file_df=df[["FDE_FECHA", "FDE_MONEDA"]])
        elif file.name.startswith("CobranzaPER"):
            # Change the values in the column "FDE_MONEDA" and "FAC_MONEDA" to "SOL"
            df["FDE_MONEDA"] = "SOL"
            df["FAC_MONEDA"] = "SOL"
            # Add missing combinations to the exchange rate file
            added_combinations = update_exchange_rate_file(date_column="FDE_FECHA", currency_column="FDE_MONEDA", file_df=df[["FDE_FECHA", "FDE_MONEDA"]])
        elif file.name.startswith("CompensacionesPAN"):
            # Change the values in the column "Moneda" to "USD"
            df["Moneda"] = "USD"
            # Add missing combinations to the exchange rate file
            added_combinations = update_exchange_rate_file(date_column="Fecha", currency_column="Moneda", file_df=df[["Fecha", "Moneda"]])
        elif file.name.startswith("CompensacionesPER"):
            # Change the values in the column "Moneda" to "SOL"
            df["Moneda"] = "SOL"
            # Add missing combinations to the exchange rate file
            added_combinations = update_exchange_rate_file(date_column="Fecha", currency_column="Moneda", file_df=df[["Fecha", "Moneda"]])
        elif file.name.startswith("NotasCrePAN"):
            # Change the values in the column "NCRE_MONEDA" to "USD"
            df["NCRE_MONEDA"] = "USD"
            # Add missing combinations to the exchange rate file
            added_combinations = update_exchange_rate_file(date_column="NCRE_FECHA", currency_column="NCRE_MONEDA", file_df=df[["NCRE_FECHA", "NCRE_MONEDA"]])
        elif file.name.startswith("NotasCrePER"):
            # Change the values in the column "NCRE_MONEDA" to "SOL"
            df["NCRE_MONEDA"] = "SOL"
            # Add missing combinations to the exchange rate file
            added_combinations = update_exchange_rate_file(date_column="NCRE_FECHA", currency_column="NCRE_MONEDA", file_df=df[["NCRE_FECHA", "NCRE_MONEDA"]])
        
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
        # Skip files in the first level of the folder structure that start with "Scripts"
        relative_path = file.relative_to(destination_folder)
        print(relative_path)
        if relative_path.parts[0].startswith("Scripts"):
            print(f"Skipping file: {file}")
            continue
            
        # Read the CSV file with the correct encoding
        df = pd.read_csv(file, encoding='ISO-8859-1', on_bad_lines='warn')
        
        # Save the Excel file
        df.to_excel(file.with_suffix('.xlsx'), index=False)

        # Remove the original CSV file
        file.unlink()  # Delete the original CSV file

        filesConverted += 1

    return filesConverted


##### Main script
# Initialize variables
files_folder = "C:/Users/pc/Desktop/ba-files/Otros" # C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/Otros     # C:/Users/pc/Desktop/ba-files/Otros
destination_folder = "C:/Users/pc/Desktop/ba-files/DYCUSA" # C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/DYCUSA     # C:/Users/pc/Desktop/ba-files/DYCUSA
logs_folder = "C:/Users/pc/Desktop/ba-files/Logs" # C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/Logs     # C:/Users/pc/Desktop/ba-files/Logs

# Create log file with current date in YYMMDD format
log_filename = f"movimientosArchivos{datetime.now().strftime('%y%m%d')}.txt"
log_path = Path(logs_folder) / log_filename

# Create logs directory if it doesn't exist
log_path.parent.mkdir(parents=True, exist_ok=True)

# Move "Otros" files to new folder
filesMoved = move_files(files_folder, destination_folder)
with open(log_path, 'a') as log_file:
    if filesMoved > 0:
        message = f"Files moved successfully: {filesMoved}\n"
        log_file.write(message)
        print(message)
    else:
        message = "No files where moved\n"
        log_file.write(message)
        print(message)

# Convert pending csv files to xlsx
extraFilesConverted = convert_pending_csv_to_xlsx(destination_folder)
with open(log_path, 'a') as log_file:
    if extraFilesConverted > 0:
        message = f"Extra files converted successfully: {extraFilesConverted}\n"
        log_file.write(message)
        print(message)
    else:
        message = "No extra files where converted\n"
        log_file.write(message)
        print(message)
