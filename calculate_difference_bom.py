# cli.py and setup.py are needed for pyinstaller package to create a distributable *.exe file

import pandas as pd
import format_input as fi 
import pathlib


excel_folders = ["v0_original_excels", "v1_updated_excels"]
df_list = []


def get_col_widths(dataframe):
    # First we find the maximum length of the index column   
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]


def format_and_write_excel(df_input):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    
    writer = pd.ExcelWriter(pathlib.Path.cwd() / "DELTA_LIST.xlsx", engine='xlsxwriter')
    df_input.to_excel(writer, sheet_name='Sheet1')
        
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    
    # Add a format. Light red fill with dark red text.
    format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                'font_color': '#9C0006'})

    worksheet = writer.sheets['Sheet1']

    # Set the conditional format range.
    start_row = 1
    start_col = 4
    end_row = len(df_input)
    end_col = start_col

    # Apply a conditional format to the cell range.
    worksheet.conditional_format(start_row, start_col, end_row, end_col,
                                {'type':'cell',
                                'criteria':'<',
                                'value':0,
                                'format':format1})

    # set column width
    for i, width in enumerate(get_col_widths(df_input)):
        worksheet.set_column(i, i, width + 2)

    # hide first index column 'Puma Code'
    worksheet.set_column('A:A', None, None, {'hidden': True})
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    

def main():
    for folder in excel_folders:
        fi.subfolder_of_excel_files = folder
        fi.file_counter_log = 0
        df = fi.concat_files(fi.get_file_list())
        df_formatted = fi.format_data(df)
        df_list.append(df_formatted)

    df_original = df_list[0]
    df_updated = df_list[1]

    # multiply original quantities by -1 to be able to add to updated values simply
    df_original["Quantity"] = -1 * df_original["Quantity"]
    # writing of intermediate excel files disabled by default. Enable my uncommenting.
    # df_original.to_excel(pathlib.Path.cwd() / "v0_original_excels" /"minus_one_output.xls", index=False)

    # generated consolidated list with original (negative values) and updated (positive values). All rows are in.
    df_delta = pd.concat([df_original, df_updated])
    # writing of intermediate excel files disabled by default. Enable my uncommenting.
    # df_delta.to_excel(pathlib.Path.cwd() / "delta_output_all_data_in.xls", index=False)

    # create sum quantities based on same Puma Code
    df_delta = df_delta.groupby(df_delta.index).agg({'Description':'first', 'NPD':'last','Puma Code':'last','Quantity':'sum'})

    # drop rows with 0 quantitiy
    df_delta = df_delta[(df_delta != 0).all(1)]

    # output writing
    #df_delta.to_excel(pathlib.Path.cwd() / "DELTA_LIST.xls", index=False)
    format_and_write_excel(df_delta)

    # wait for user to close terminal
    prompt = "Excel compare finished. Hit any key to exit. See log.log file for results."
    user_input = input(prompt)


if __name__ == "__main__":
    main()