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


def process_preferential_sheet(
        xl,
        preferential_sheet_name,
        country_name,
        rta_name,
        comtrade_country_mapping,
        economic_blocks_list,
):
    try:
        # Parse everything as strings (this may be redundant)
        columns = xl.parse(sheet_name=preferential_sheet_name, header=[0, 1]).columns
        converters = {column: str for column in columns}
        preferential_df = xl.parse(sheet_name=preferential_sheet_name, header=[0, 1], converters=converters)

        # Remove spurious lines
        preferential_df = preferential_df.dropna(how='all')

        columns_to_keep = []
        for column in preferential_df.columns.get_level_values(0).array:
            column = str(column)
            if column in ["TL", "TLS", "Year", "Reporter"] and column not in columns_to_keep:
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

        temp_df = preferential_df.melt(col_level=1, id_vars=id_vars, var_name="Tariff_Year", value_name="Tariff")
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
        # Parse everything as strings (this may be redundant)
        columns = xl.parse(sheet_name=preferential_sheet_name).columns
        converters = {column: str for column in columns}
        preferential_df = xl.parse(sheet_name=preferential_sheet_name, converters=converters)

        # Remove spurious lines
        preferential_df = preferential_df.dropna(how='all')

        columns_to_keep = []
        for column in preferential_df.columns.get_level_values(0).array:
            column = str(column)
            if column in ["TL", "TLS", "Year", "Reporter"] and column not in columns_to_keep:
                columns_to_keep.append(column)
            if "Preferential" in column and column not in columns_to_keep:
                preferential_column_rename = {column: "Tariff"}
                columns_to_keep.append(column)

        preferential_df = preferential_df[columns_to_keep]
        preferential_df["Tariff_Year"] = preferential_df["Year"]

        preferential_df.rename(columns=preferential_column_rename, inplace=True)

    except:
        raise ParsePreferentialError

    # Identify reporter country
    reporter_country = None
    reporter_country_code = None

    # Economic block identification should have the least priority among other options since some
    # agreements differentiate tariffs between countries in the block (e.g. some EFTA related agreements)
    if "Reporter" in preferential_df.columns:
        reporter_identifier = preferential_df["Reporter"].iloc[0]
        # Identifier is a name
        if reporter_identifier in comtrade_country_mapping["Country Name, Full "].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Name, Full "] == reporter_identifier]
            reporter_country = reporter_identifier
            reporter_country_code = row["Country Code"].iloc[0]
        elif reporter_identifier in comtrade_country_mapping["Country Name, Abbreviation"].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Name, Abbreviation"] == reporter_identifier]
            reporter_country = row["Country Name, Full "].iloc[0]
            reporter_country_code = row["Country Code"].iloc[0]
        elif reporter_identifier in comtrade_country_mapping["Country Name, Other"].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Name, Other"] == reporter_identifier]
            reporter_country = row["Country Name, Full "].iloc[0]
            reporter_country_code = row["Country Code"].iloc[0]
        elif reporter_identifier in comtrade_country_mapping["Country Name, Other Abbreviation"].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Name, Other Abbreviation"] == reporter_identifier]
            reporter_country = row["Country Name, Full "].iloc[0]
            reporter_country_code = row["Country Code"].iloc[0]
        elif reporter_identifier in economic_blocks_list:
            reporter_country = reporter_identifier
        # Identifier is a code
        elif int(reporter_identifier) in comtrade_country_mapping["Country Code"].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Code"] == int(reporter_identifier)]
            reporter_country = row["Country Name, Full "].iloc[0]
            reporter_country_code = int(reporter_identifier)
        # Use country name as reporter
        elif country_name in economic_blocks_list:
            reporter_country = country_name
    else:
        raise ParsePreferentialError

    # Identify partner country
    partner_country = None
    partner_country_code = None
    partner_candidates = []
    # Source candidates from preferential sheet name - more discriminative in corrected data
    for partner_candidate in preferential_sheet_name.split("_"):
        partner_candidates.append(partner_candidate.strip(" "))
    # Source candidates form RTA name
    partner_candidates.append(rta_name)
    for partner_candidate in rta_name.split("-"):
        partner_candidates.append(partner_candidate.strip(" "))

    for partner_candidate in partner_candidates:
        if partner_candidate == country_name or partner_candidate == reporter_country:
            continue
        elif partner_candidate in comtrade_country_mapping["Country Name, Full "].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Name, Full "] == partner_candidate]
            partner_country = partner_candidate
            partner_country_code = row["Country Code"].iloc[0]
            break
        elif partner_candidate in comtrade_country_mapping["Country Name, Abbreviation"].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Name, Abbreviation"] == partner_candidate]
            partner_country = row["Country Name, Full "].iloc[0]
            partner_country_code = row["Country Code"].iloc[0]
            break
        elif partner_candidate in comtrade_country_mapping["Country Name, Other"].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Name, Other"] == partner_candidate]
            partner_country = row["Country Name, Full "].iloc[0]
            partner_country_code = row["Country Code"].iloc[0]
            break
        elif partner_candidate in comtrade_country_mapping["Country Name, Other Abbreviation"].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["Country Name, Other Abbreviation"] == partner_candidate]
            partner_country = row["Country Name, Full "].iloc[0]
            partner_country_code = row["Country Code"].iloc[0]
            break
        elif partner_candidate in comtrade_country_mapping["ISO3-digit Alpha"].values:
            row = comtrade_country_mapping.loc[comtrade_country_mapping["ISO3-digit Alpha"] == partner_candidate]
            partner_country = row["Country Name, Full "].iloc[0]
            partner_country_code = row["Country Code"].iloc[0]
            break

    # Economic block identification should have the least priority among other options since some
    # agreements differentiate tariffs between countries in the block (e.g. some EFTA related agreements)
    if partner_country is None:
        for partner_candidate in partner_candidates:
            if partner_candidate == country_name or partner_candidate == reporter_country:
                continue
            elif partner_candidate in economic_blocks_list:
                partner_country = partner_candidate
                break

    if partner_country is None:
        raise ParsePreferentialError

    preferential_df = preferential_df.assign(Reporter=reporter_country)
    preferential_df = preferential_df.assign(Reporter_Code=reporter_country_code)
    preferential_df = preferential_df.assign(Partner=partner_country)
    preferential_df = preferential_df.assign(Partner_Code=partner_country_code)
    preferential_df = preferential_df.assign(RTA=rta_name)

    return preferential_df


def process_file(file_name):
    print(f'Processing file: {file_name}')

    df = pd.DataFrame()

    xl = pd.ExcelFile(file_name)

    preferential_sheet_names = []
    for x in xl.sheet_names:
        if "Preferential" in x:
            preferential_sheet_names.append(x)
            continue
        columns = xl.parse(sheet_name=x).columns
        for column in columns:
            if "Preferential" in column:
                preferential_sheet_names.append(x)

    # Do not silently fail if no preferential sheets are found
    if len(preferential_sheet_names) == 0:
        raise ParsePreferentialError

    # Parse Notes sheet - assumes its the first one
    notes_df = xl.parse(
        sheet_name=0,
        header=None,
        usecols=[0, 1, 2],
        names=["Label", "Colon", "Value"],
        engine="openpyxl"
    )

    rta_name = notes_df.loc[notes_df["Label"].isin(["RTA", "FTA", "Agreement", "PTA"])]["Value"].to_list()[0]
    country_name = notes_df.loc[notes_df["Label"].isin(["Country"])]["Value"].to_list()[0]

    # Parse Preferential Rates Sheet
    for preferential_sheet_name in preferential_sheet_names:
        preferential_df = process_preferential_sheet(
            xl,
            preferential_sheet_name,
            country_name,
            rta_name,
            comtrade_country_mapping,
            economic_blocks_list
        )
        df = df.append(preferential_df)

    # At least one sheet containing preferential info should have been found
    if len(df) == 0:
        raise ParsePreferentialError

    return df


def search_files():

    df = pd.DataFrame()

    # Blacklist for files not yet supported or out of scope for Preferential Tariffs
    file_blacklist = [
        # EFTA-Chile RTA data - old and inconsistent format
        "EFTA_3.xls",
        "Chile_11.xls",
        # CET agreement of EAEU
        "Armenia_2015.xls", "Armenia_2016.xls",
        "Belarus2015.xls", "Belarus2016.xls",
        "Kyrgyz Republic_2015.xls", "Kyrgyz Republic_2016.xls",
        "Kazakhstan2015.xls", "Kazakhstan2016.xls",
        "Russian Federation2015.xls", "Russian Federation2016.xls",
        # Commerce blocks' related files
        "SACU_1.xls",
    ]

    economic_blocks_list = [
        "EU",  # European Union
        "European Union",
        "SACU",  # Southern African Customs Union
        "Trans-Pacific SEP",  # Trans-Pacific Strategic Economic Partnership Agreement
        "CAFTA",  # Central America Free Trade Agreement
        "Agadir Agreement",
        "EFTA",
        "EAEU",  # Eurasian Economic Union
        "MERCOSUR",
        "ASEAN",
    ]

    path = Path(__file__).parent.parent / 'data'
    comtrade_country_mapping = pd.read_excel(path / "Comtrade Country Code and ISO list.xlsx")

    already_parsed_files = []
    for dirpath, dirnames, files in os.walk(path / "RTAs"):
        print(f'Found directory: {dirpath}')
        for file_name in files:
            if file_name in file_blacklist or file_name in already_parsed_files:
                continue
            try:
                full_file_name = str(path / dirpath / file_name)

                # For debugging
                """
                reporter_country = os.path.basename(os.path.dirname(full_file_name))
                if reporter_country < "Japan":
                    continue

                if file_name == "MERCOSUR.xls":
                    a = 1
                """
                temp_df = process_file(full_file_name, comtrade_country_mapping, economic_blocks_list)
                df = df.append(temp_df)
                already_parsed_files.append(file_name)
            except:
                df.to_csv("partial_rta_file.csv", index=False)
                raise ParseFileError
    df.to_csv("full_rta_file.csv", index=False)

search_files()
