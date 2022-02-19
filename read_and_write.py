import pandas as pd
import json

def read_json(jsonPath, dict=''):
    """ Lee json con la información """
    user = []
    with open(jsonPath) as c:
        login = json.load(c)[dict]
        [user.append(i) for i in login]
    return user[0]

def substitute(path):
    """ Convierte el csv a xlsx en una tabla """
    pathxlsx = read_json('config.json', 'login')['pathd']
    pd.read_csv(path).to_excel(pathxlsx, index=None, header=True)
    xls = pd.ExcelFile(pathxlsx)
    df1 = xls.parse(xls.sheet_names[0])
    df = pd.DataFrame(df1.to_dict())
    writer = pd.ExcelWriter(pathxlsx, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1',
                startrow=1, header=False, index=False)
    worksheet = writer.sheets['Sheet1']
    (max_row, max_col) = df.shape
    column_settings = []
    for header in df.columns:
        column_settings.append({'header': header})
    worksheet.add_table(0, 0, max_row, max_col - 1,
                        {'columns': column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    del df, df1, column_settings
    writer.save()   