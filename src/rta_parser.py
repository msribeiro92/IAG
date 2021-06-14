import os
from pathlib import Path

import numpy as np
import win32com.client as win32
import pandas as pd

from src.parse_file_error import ParseFileError
from src.parse_preferential_error import ParsePreferentialError


def convert_file(file_name):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(file_name)

    wb.SaveAs(file_name + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
    wb.Close()                             # FileFormat = 56 is for .xls extension
    excel.Application.Quit()


def process_preferential_sheet(xl, preferential_sheet_name, reporter_country, rta_name):
    try:
        preferential_df = xl.parse(sheet_name=preferential_sheet_name, header=[0, 1])

        columns_to_keep = []
        for column in preferential_df.columns.get_level_values(0).array:
            column = str(column)
            if column in ["TL", "TLS", "Year"] and column not in columns_to_keep:
                columns_to_keep.append(column)
            if "Preferential" in column and column not in columns_to_keep:
                columns_to_keep.append(column)

        preferential_df = preferential_df[columns_to_keep]

        # Fixes multiple preferential columns spuriously created when
        # multi-index columns are irrelevant
        preferential_df = preferential_df.replace("*", np.nan)
        preferential_df = preferential_df.dropna(axis='columns', how='all')

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

        temp_df = preferential_df.melt(col_level=1, id_vars=id_vars, var_name="Tariff Year", value_name="Tariff")
        if not temp_df.empty:
            preferential_df = temp_df
            preferential_df.rename(columns=id_vars_to_names, inplace=True)
        # Fixes multiple preferential columns spuriously created when
        # multi-index columns are irrelevant
        else:
            preferential_df = preferential_df.droplevel(1, axis='columns')
            preferential_df = preferential_df.rename(columns=lambda x: "Tariff" if "Preferential" in x else x)
            preferential_df["Tariff_Year"] = preferential_df["Year"]

    except pd.errors.ParserError:
        preferential_df = xl.parse(sheet_name=preferential_sheet_name)

        columns_to_keep = []
        for column in preferential_df.columns.get_level_values(0).array:
            column = str(column)
            if column in ["TL", "TLS", "Year"] and column not in columns_to_keep:
                columns_to_keep.append(column)
            if "Preferential" in column and column not in columns_to_keep:
                preferential_column_rename = {column: "Tariff"}
                columns_to_keep.append(column)

        preferential_df = preferential_df[columns_to_keep]
        preferential_df["Tariff_Year"] = preferential_df["Year"]

        preferential_df.rename(columns=preferential_column_rename, inplace=True)

    except:
        raise ParsePreferentialError

    preferential_df = preferential_df.assign(Type="Preferential")
    preferential_df = preferential_df.assign(Country=reporter_country)
    preferential_df = preferential_df.assign(RTA=rta_name)
    preferential_df = preferential_df.assign(RTA_Detail=preferential_sheet_name)

    return preferential_df


def process_file(file_name):
    print(f'Processing file: {file_name}')

    df = pd.DataFrame()

    reporter_country = os.path.basename(os.path.dirname(file_name))

    xl = pd.ExcelFile(file_name)

    preferential_sheet_names = []
    for x in xl.sheet_names:
        if "Preferential" in x:
            preferential_sheet_names.append(x)

    # Skip if no Preferential Information
    if len(preferential_sheet_names) == 0:
        return df

    # Parse Notes sheet - assumes its the first one
    notes_df = xl.parse(
        sheet_name=0,
        header=None,
        usecols=[0, 1, 2],
        names=["Label", "Colon", "Value"],
        engine="openpyxl"
    )

    rta_name = notes_df.loc[notes_df["Label"].isin(["RTA", "FTA", "Agreement", "PTA"])]["Value"].to_list()[0]

    # Parse Preferential Rates Sheet
    for preferential_sheet_name in preferential_sheet_names:
        preferential_df = process_preferential_sheet(
            xl,
            preferential_sheet_name,
            reporter_country,
            rta_name
        )
        df = df.append(preferential_df)

    return df


def search_files():
    df = pd.DataFrame()

    # Blacklist for files not yet supported
    # EFTA-Chile RTA data - old and inconsistent format
    file_blacklist = ["EFTA_3.xls"]

    path = Path(__file__).parent.parent / 'data'
    for dirpath, dirnames, files in os.walk(path):
        print(f'Found directory: {dirpath}')
        for file_name in files:
            if file_name in file_blacklist:
                continue
            try:
                full_file_name = str(path / dirpath / file_name)
                """
                # For debugging
                reporter_country = os.path.basename(os.path.dirname(full_file_name))
                if reporter_country < "Iceland":
                    continue
                if file_name == "Turkey_rev1.xls":
                    a = 1
                """
                temp_df = process_file(full_file_name)
                df = df.append(temp_df)
            except:
                df.to_csv("partial_rta_file.csv")
                raise ParseFileError
    df.to_csv("full_rta_file.csv")


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
