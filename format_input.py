import pandas as pd
import pathlib
import logging


subfolder_of_excel_files = ""
logging.basicConfig(format='%(asctime)s - %(message)s', datefmt='%d-%m-%y %H:%M:%S', filename='log.log', level=logging.DEBUG)
logging.info('Started')


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
    global file_counter_log 
    current_file = file_list[0]
    try:
        df = concat_sheets(current_file)
        file_counter_log += 1
        logging.info(f"file: {current_file.name} processed")
    except:
        logging.info(f"NOT PROCESSED file: {current_file.name}")
    
    file_list.pop(0)
    for current_file in file_list:
        try:
            df_current_file = concat_sheets(current_file)
            df = pd.concat([df, df_current_file])
            file_counter_log += 1
            logging.debug(f"file: {current_file.name} processed")
        except:
            logging.info(f"NOT PROCESSED file: {current_file.name}")
        
    logging.debug(f"Number of files processed: {file_counter_log}")
    return df


def format_data(df):
    # drop first column which only contains Q1 for all lines (SP3D macro)
    # drop "Item number" column, which is also not needed
    df.drop(df.columns[0], axis=1, inplace=True)
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
    df_formatted = df.groupby(['Puma Code']).agg({'Description':'first', 'NPD':'last','Puma Code':'last','Quantity':'sum'})
    
    # writing of intermediate excel files disabled by default. Enable my uncommenting.
    # df_formatted.to_excel(pathlib.Path.cwd() / subfolder_of_excel_files /"processed_output.xls", index=False)
    return df_formatted


# subfolder_of_excel_files = "v0_original_excels"
# file_counter_log = 0
# df = concat_files(get_file_list())
# format_data(df)
