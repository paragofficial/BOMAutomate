# api_functions.py

import requests
import pandas as pd

def make_api_request(url, headers, data):
    response = requests.post(url, headers=headers, json=data)
    return response

def construct_query_template(mpn_values):
    query_template = '''
    query MultiSearch {
        supMultiMatch (
            queries: [
                %s
            ]
        ){
            hits
            parts {
                id
                name
                mpn
            }
        }
    }
    '''

    queries = ',\n'.join('{{mpn: "{}"}}'.format(mpn) for mpn in mpn_values)
    final_query = query_template % queries
    return final_query

def process_api_response(result, mpn_values):
    df_result = pd.DataFrame({'MPN': mpn_values, 'Exists': 'No'})

    for supMultiMatch in result.get('data', {}).get('supMultiMatch', []):
        for part in supMultiMatch.get('parts', []):
            mpn = part.get('mpn', '')
            if mpn in mpn_values:
                df_result.loc[df_result['MPN'] == mpn, 'Exists'] = 'Yes'

    print(df_result)
    return df_result
