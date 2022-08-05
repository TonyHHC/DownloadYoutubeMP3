import os
import glob
import pandas as pd
import re
import openpyxl
import uuid

strLine1 = '{"type": "HDT", "version": "1.2"}'
strLineN = '{"ver": "1.2", "title": "Waiting... $<No>", "gal_num": "$<youtube_url>", "music": false, "anime": true, "valid": true, "dir": "", "label_color": null, "url": "$<youtube_url>", "type": "youtube", "filesize": 0, "pbar": [0, 0, "%v/%m"], "time": 1659516032.9042757, "version": "3.7h", "uid": "$<uuid>", "str_pixmap": "", "artist": null, "done": false, "name_zip": "", "urls": [], "etc_button": null, "pad": true}'

def getUrls(strExcelName, strSheetName):
    df_Spec = pd.DataFrame(columns=['No', 'url', 'uuid'])

    workbook = openpyxl.load_workbook(strExcelName, data_only=True)
    st = workbook[strSheetName]
    
    print('Sheet : ' + strSheetName)
    
    for row in st.iter_rows(min_row=0):
        if len(row) > 3 :
            if str(row[3].value).startswith('https'):
                strUUID = uuid.uuid4()
                df_Spec = df_Spec.append({'No':str(int(row[0].value)), 'url':str(row[3].value), 'uuid':str(uuid.uuid4()).replace('-','')}, ignore_index=True)
            
    return df_Spec
    
if __name__ == "__main__":

    with open('billboard.hdt', mode='w', encoding='utf-8') as w:
        w.write('[')
        w.write('\n' + strLine1)

        for iYear in range(1960, 2003):
            df_Spec = getUrls('Billboards Top 100.xlsx', str(iYear))
            print(df_Spec)
            for index, row in df_Spec.iterrows():
                strLine = strLineN
                for r in (('$<No>', row['No']), ('$<youtube_url>', row['url']), ('$<uuid>', row['uuid'])):
                    strLine = strLine.replace(*r)
                
                w.write(',\n' + strLine)
                    
        w.write('\n]')