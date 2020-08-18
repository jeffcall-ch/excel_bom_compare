import pandas as pd
import format_input as fi 

excel_file = "AD00046190.xls"

origi_file = fi.FormatInput(excel_file)

df = origi_file.get_df()
df.to_excel(excel_file[:-4]+"_output.xls", index=False)
