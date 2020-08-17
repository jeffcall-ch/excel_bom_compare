import pandas as pd


pd.set_option('display.max_columns', 10)  # or 1000
pd.set_option('display.max_rows', None)  # or 1000
pd.set_option('display.max_colwidth', 20)  # or 199

excel_file = 'AD00046190.xls'
pd_xls_list = []
bom_data = None


def get_visible_sheet_names():
    xls_file = pd.ExcelFile(excel_file)
    sheets = xls_file.book.sheets()
    sheet_list = []
    for sheet in sheets:
        if sheet.visibility == 0:
            sheet_list.append(sheet.name)
    return sheet_list


def format_data():
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


visible_sheets = get_visible_sheet_names()
#concat sheets
file = pd.ExcelFile(excel_file)
df = pd.concat([file.parse(sheet_name, skiprows=3) for sheet_name in visible_sheets])

format_data()


# create sum quantities based on same Puma Code
df1 = df.groupby(['Puma Code']).agg({'Description':'first', 'NPD':'last','Puma Code':'last','Quantity':'sum'})

df1.to_excel('output.xls', index=False)



