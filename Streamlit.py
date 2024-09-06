import os
import pandas as pd
import numpy as np
import streamlit as st
from datetime import date

def import_excel_files(directory_path, tarieven_file_path, directory_path_output):
    all_data = pd.DataFrame()
    input_data = pd.read_excel(tarieven_file_path, sheet_name="Tarieven", usecols='A:C')
    input_data['Klantnummer'] = input_data['Klantnummer'].astype(str)
    klantnaam_list = input_data['Klantnummer'].values.tolist()
    klantnaam_list = map(str, klantnaam_list)

    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith(".xlsx") or file.endswith(".xls"):
                file_path = os.path.join(root, file)
                excel_data = pd.read_excel(file_path)
                all_data = all_data.append(excel_data)

    all_data = all_data.dropna(how="all")
    all_data.columns = all_data.iloc[1]
    all_data = all_data[1:]
    all_data = all_data.rename(columns={'Klant nummer': 'Klantnummer'})

    if 'Klantnummer' in all_data.columns:
        all_data['Klantnummer'] = all_data['Klantnummer'].astype(str)
    else:
        st.error("Error: 'Klantnummer' column not found in all_data")
        return

    all_data["Tarief"] = 0
    all_data["Klantnaam"] = "leeg"
    all_data["Aantal_colli"] = 0
    all_data["land"] = "NL"
    all_data['Order nr. verlader'] = all_data['Order nr. verlader'].astype(str)
    all_data['Verlader naam'] = all_data['Verlader naam'].astype(str)
    all_data = all_data.reset_index(drop=True)

    count_colli = all_data.groupby('Order nr. verlader', as_index=False).size()
    
    for k in klantnaam_list:
        for i in range(len(all_data)):
            if all_data["Klantnummer"][i] == k:
                rownumber = int(input_data[input_data['Klantnummer'] == k].index[0])
                tarief = input_data.loc[rownumber, "Tarief"]
                klantnaam = input_data.loc[rownumber, "Klantnaam"]
                ordernrverlader = all_data["Order nr. verlader"][i]
                colli = (count_colli.loc[(count_colli['Order nr. verlader'] == ordernrverlader)].values[0])[1]
                all_data.at[i, "Tarief"] = tarief
                all_data.at[i, "Klantnaam"] = klantnaam
                all_data.at[i, "Aantal_colli"] = colli
                all_data.at[i, "Order nr. verlader"] = ordernrverlader

    all_data.drop_duplicates(subset=['Order nr. verlader'], keep="first", inplace=True)
    all_data = all_data.reset_index(drop=True)
    
    columns_list = ['index', 'level_0']
    all_data = all_data.drop(columns=columns_list, errors='ignore')
    
    pivot_table = all_data.pivot_table(index=['Klantnummer', 'Klantnaam'], values=['Tarief'], aggfunc=['count', 'sum'])
    O2C_data = all_data[['Klantnaam', 'Order nr. verlader', 'Tarief', 'Klantnummer', 'land', 'Aantal_colli']]
    O2C_data['productcode'] = 3762
    O2C_data['datum'] = date.today()
    
    new_index = ['Klantnummer', 'Klantnaam', 'productcode', 'land', 'Order nr. verlader', 'Tarief', 'datum', 'Aantal_colli']
    O2C_data = O2C_data.reindex(columns=new_index).iloc[1:, :].replace('nan', np.nan).dropna(subset=['Klantnaam'])

    with pd.ExcelWriter(directory_path_output) as writer:
        all_data.to_excel(writer, sheet_name="Combined_data")
        pivot_table.to_excel(writer, sheet_name="Pivot_Table")
        O2C_data.to_excel(writer, sheet_name="Input voor O2C")

    return input_data, all_data

st.title("EH retouren app")
st.title("Welcome")
print("check")

input_directory_path = st.text_input("Input Directory Path")
tarieven_file = st.file_uploader("Upload Tarieven File", type=["xlsx", "xls"])
output_directory_path = st.text_input("Output Directory Path")

if st.button("Execute"):
    if input_directory_path and tarieven_file and output_directory_path:
        directory_path_output = os.path.join(output_directory_path, "output_combined_data.xlsx")
        combined_data = import_excel_files(input_directory_path, tarieven_file, directory_path_output)
        st.success("Script executed successfully. New Excel file created.")
    else:
        st.error("Please provide input directory, tarieven file, and output directory.")

