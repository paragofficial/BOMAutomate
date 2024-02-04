# excel_function.py

import pandas as pd
from api_function import construct_query_template, make_api_request, process_api_response

def process_excel_data_and_save(excel_file_path, sheet_name, url, headers):
    df_excel = read_excel_data(excel_file_path, sheet_name)
    mpn_values = extract_mpn_values(df_excel)

    query = construct_query_template(mpn_values)
    data = {'query': query}

    response = make_api_request(url, headers, data)

    if response.status_code == 200:
        result = response.json()
        df_result = process_api_response(result, mpn_values)
        df_result.to_excel('exists.xlsx', index=False)
        print("Result saved to exists.xlsx")
    else:
        print(f"Error: {response.status_code}, {response.text}")

def read_excel_data(excel_file_path, sheet_name):
    df_excel = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    return df_excel

def extract_mpn_values(df_excel):
    mpn_values = df_excel['MPN'].tolist()
    return mpn_values
