import pandas as pd
import numpy as np

no_3s_excel = pd.read_excel('Uniform_inventory_ccf.xlsx', sheet_name='No. 3s')
PCS_excel = pd.read_excel('Uniform_inventory_ccf.xlsx', sheet_name='PCS')
foul_weather_excel = pd.read_excel('Uniform_inventory_ccf.xlsx', sheet_name='Foul Weather Jacket')
headdress_excel = pd.read_excel('Uniform_inventory_ccf.xlsx', sheet_name='Headdress')
epaulette_excel = pd.read_excel('Uniform_inventory_ccf.xlsx', sheet_name='Epaulettes')
other_excel = pd.read_excel('Uniform_inventory_ccf.xlsx', sheet_name='Other')

def sheet_and_column_selector(sheet_name, col_name):
    if sheet_name == 'No. 3s':
        df = pd.DataFrame(no_3s_excel)
    elif sheet_name == 'PCS':
        df = pd.DataFrame(PCS_excel)
    elif sheet_name == 'Foul Weather Jacket':
        df = pd.DataFrame(foul_weather_excel)
    elif sheet_name == 'Headdress':
        df = pd.DataFrame(headdress_excel)
    elif sheet_name == 'Epaulettes':
        df = pd.DataFrame(epaulette_excel)
    elif sheet_name == 'Other':
        df = pd.DataFrame(other_excel)
    else:
        raise ValueError("Invalid sheet name provided.")
    df = df[[col_name]].copy()
    return df
def insert_data_to_excel(sheet_name, col_name, size_name, data_input):
    df = sheet_and_column_selector(sheet_name, col_name)
    target_column = df.columns.get_loc(col_name) + 1
    df.at[size_name, df.columns[target_column]] = data_input
    with pd.ExcelWriter('Uniform_inventory_ccf.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

insert_data_to_excel('No. 3s', 'Male Shirts', 32, 10)