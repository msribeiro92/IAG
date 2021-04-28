import os
from pathlib import Path
import win32com.client as win32
import pandas as pd


def convert_file(file_name):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(file_name)

    wb.SaveAs(file_name + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
    wb.Close()                             # FileFormat = 56 is for .xls extension
    excel.Application.Quit()


def process_file(file_name):
    print(f'Processing file: {file_name}')
    xl = pd.ExcelFile(file_name)

    # Parse Notes sheet - assumes its the first one
    notes_df = xl.parse(sheet_name=0, header=None, names=["Label", "Colon", "Value"])
    reporter_country = notes_df.loc[notes_df["Label"] == "Country"]["Value"].to_list()[0]
    rta_name = notes_df.loc[notes_df["Label"].isin(["RTA", "FTA"])]["Value"].to_list()[0]

    # Parse Preferential Rates Sheet
    preferential_sheet_name = next(x for x in xl.sheet_names if "Preferential" in x)
    preferential_df = xl.parse(sheet_name=preferential_sheet_name, header=[0, 1])

    columns_to_keep = []
    for column in preferential_df.columns.get_level_values(0).array:
        column = str(column)
        if column in ["TL"] and column not in columns_to_keep:
            columns_to_keep.append(column)
        if "Preferential" in column and column not in columns_to_keep:
            columns_to_keep.append(column)

    preferential_df = preferential_df[columns_to_keep]

    # Fix MultiIndex column names
    column_0 = preferential_df.columns.get_level_values(0).array
    column_1 = preferential_df.columns.get_level_values(1).array
    id_vars_to_names = {}
    id_vars = []
    for i, column in enumerate(column_1):
        column = str(column)
        if "Unnamed" in column:
            id_vars.append(column)
            id_vars_to_names[column] = column_0[i]

    preferential_df = preferential_df.melt(col_level=1, id_vars=id_vars, var_name="Year", value_name="Tariff")
    preferential_df.rename(columns=id_vars_to_names, inplace=True)
    preferential_df = preferential_df.assign(Type="Preferential")
    preferential_df = preferential_df.assign(Country=reporter_country)
    preferential_df = preferential_df.assign(RTA=rta_name)

    return preferential_df


def search_files():
    df = pd.DataFrame()

    path = Path(__file__).parent.parent / 'data'
    for dirpath, dirnames, files in os.walk(path):
        print(f'Found directory: {dirpath}')
        for file_name in files:
            full_file_name = str(path / dirpath / file_name)
            temp_df = process_file(full_file_name)
            df = df.append(temp_df)


def init_data_frame():
    schema = {
        'Reporter': [],
        'Year': [],
        'Tariff Type': [],
        'Code': [],
        'Tariff': [],
        'RTA Name': []
    }
    df = pd.DataFrame(schema)

    return df


search_files()