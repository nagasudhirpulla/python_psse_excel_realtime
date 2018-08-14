# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import requests
import json

def getRTData(pnt):
    # print(req_date_str)
    params = dict(
            pnt=pnt
            )
    r = requests.get(url = "http://wmrm0mc1:62448/api/values/real", params = params)
    data = json.loads(r.text)
    return data


# handle WRGEN_values sheet
df = pd.read_excel('input.xlsx', sheetname='WRGEN_values')

# iterate through all rows
for k in range(df.shape[0]):
    # find the real time value
    rtData = getRTData(df.iloc[k,3])
    if(rtData==None):
        continue
    df.iloc[k,4] = rtData['dval']
    # print(str(rtData['dval']) + "   " + rtData['status'])
    df.iloc[k,5] = rtData['status']
    
# handle State_load sheet
df1 = pd.read_excel('input.xlsx', sheetname='State_load')

# iterate through all rows
for k in range(df1.shape[0]):
    # find the real time value
    rtData = getRTData(df1.iloc[k,2])
    if(rtData==None):
        continue
    df1.iloc[k,1] = rtData['dval']
    
# handle HVDC sheet
df2 = pd.read_excel('input.xlsx', sheetname='HVDC')

# iterate through all rows
for k in range(df2.shape[0]):
    # find the real time value
    rtData = getRTData(df2.iloc[k,4])
    if(rtData==None):
        continue
    df2.iloc[k,2] = rtData['dval']

writer = pd.ExcelWriter('output.xlsx')   
df.to_excel(writer,sheet_name='wrgen')   
df1.to_excel(writer,sheet_name='stateload')
df2.to_excel(writer,sheet_name='hvdc')
writer.save()
