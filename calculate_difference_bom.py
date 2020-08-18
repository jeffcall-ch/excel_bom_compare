import pandas as pd
import format_input as fi 
import pathlib

excel_folders = ["v0_original_excels", "v1_updated_excels"]
df_list = []

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
df_original.to_excel(pathlib.Path.cwd() / "v0_original_excels" /"minus_one_output.xls", index=False)

# generated consolidated list with original (negative values) and updated (positive values). All rows are in.
df_delta = pd.concat([df_original, df_updated])
df_delta.to_excel(pathlib.Path.cwd() / "delta_output_all_data_in.xls", index=False)

# create sum quantities based on same Puma Code
df_delta = df_delta.groupby(df_delta.index).agg({'Description':'first', 'NPD':'last','Puma Code':'last','Quantity':'sum'})

# drop rows with 0 quantitiy
df_delta = df_delta[(df_delta != 0).all(1)]

# output writing
df_delta.to_excel(pathlib.Path.cwd() / "DELTA_LIST.xls", index=False)