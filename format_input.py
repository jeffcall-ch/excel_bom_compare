import pandas as pd
import pathlib


subfolder_of_excel_files = ""


def get_file_list():
    path_of_excel_files = pathlib.Path.cwd() / subfolder_of_excel_files
    file_list = []
    for path in path_of_excel_files.glob("*.xls"):
        file_list.append(path)
    return file_list


def get_visible_sheet_names(excel_file):
    xls_file = pd.ExcelFile(excel_file)
    sheets = xls_file.book.sheets()
    sheet_list = []
    for sheet in sheets:
        if sheet.visibility == 0:
            sheet_list.append(sheet.name)
    return sheet_list


def concat_sheets(excel_file):
    visible_sheets = get_visible_sheet_names(excel_file)
    file = pd.ExcelFile(excel_file)
    df_one_file = pd.concat([file.parse(sheet_name, skiprows=3) for sheet_name in visible_sheets])
    return df_one_file


def concat_files(file_list):
    df = concat_sheets(file_list[0])
    file_list.pop(0)
    for current_file in file_list:
        df_current_file = concat_sheets(current_file)
        df.append(df_current_file)
    df.to_excel(pathlib.Path.cwd() / "vx_delta_excel" / "inmediate_output_merged_files.xls", index=False)    
    return df


def format_data(df):
        # drop first column which only contains Q1 for all lines (SP3D macro)
        df.drop(df.columns[0], axis=1, inplace=True)
        # drop all lines which are empty
        df.dropna(subset=['Description'], inplace=True)
        # remove "m" unit from meter line items
        df['Quantity'] = df['Quantity'].replace(' m','',regex=True)
        # convert quantity to number format 
        df['Quantity'] = df['Quantity'].apply(pd.to_numeric, errors='ignore')
        # strip whitespace from Puma code
        df['Puma Code'] = df['Puma Code'].str.strip()
        # create sum quantities based on same Puma
        df = df.groupby(['Puma Code']).agg({'Description':'first', 'NPD':'last','Puma Code':'last','Quantity':'sum'})
        # df.to_excel(excel_file[:-4]+"_output.xls", index=False)
        df.to_excel(pathlib.Path.cwd() / "vx_delta_excel" / "output.xls", index=False)


subfolder_of_excel_files = "v0_original_excels"
for file in get_file_list():
    print(file)

df = concat_files(get_file_list())
format_data(df)
