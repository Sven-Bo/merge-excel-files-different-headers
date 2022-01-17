from pathlib import Path

import pandas as pd
import xlwings as xw

INPUT_DIR = Path.cwd() / "INPUT"

files = list(INPUT_DIR.glob("*.xls*"))

# Mapping table for different header names
mapping_table = {
    "Nombre": "Name",
    "Salario": "Salary",
    "Departmento": "Department",
    "Nome": "Name",
    "Stipendio": "Salary",
    "Dipartimento": "Department",
    "Abteilung": "Department",
    "Gehalt": "Salary",
}

dataframes = []

with xw.App(visible=False) as app:
    for file in files:
        # Open each excel file
        wb = app.books.open(file)
        sht = wb.sheets[0]

        # Convert table data to dataframe, rename headers, add source column & append to dataframes list
        df = sht.tables["tSalary"].range.options(pd.DataFrame, index=False).value
        df = df.rename(columns=mapping_table)
        df["Source Name"] = file.stem
        dataframes.append(df)
        wb.close()

    # Concatenate dataframes and export to new excel workbook
    combined_df = pd.concat(dataframes)
    wb_combined = app.books.add()
    wb_combined.sheets[0].range("A1").options(
        pd.DataFrame, index=False
    ).value = combined_df
    wb_combined.save(Path.cwd() / "Combined_Data.xlsx")
    wb_combined.close()
