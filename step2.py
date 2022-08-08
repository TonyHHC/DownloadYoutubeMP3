import os
from os import listdir
from os.path import isfile, isdir, join
import json
import pandas as pd
import openpyxl
import shutil
import eyed3	

df_Spec = pd.DataFrame(columns=['Sheet', 'No', 'Title', 'Artist', 'url', 'htFilename'])
df_Physical = pd.DataFrame(columns=['url', 'Filename'])


def getUrls(strExcelName, strSheetName):
	global df_Spec

	workbook = openpyxl.load_workbook(strExcelName, data_only=True)
	st = workbook[strSheetName]

	print('Sheet : ' + strSheetName)

	for row in st.iter_rows(min_row=0):
		if len(row) > 3 :
			if str(row[3].value).startswith('https'):
				df_Spec = df_Spec.append({'Sheet':strSheetName, 'No':str(int(row[0].value)).zfill(3), 'Title':str(row[1].value), 'Artist':str(row[2].value), 'url':str(row[3].value)}, ignore_index=True)
			
	
def getPhysical(strHDT_Filename):
	global df_Physical

	with open(strHDT_Filename) as f:
		datas = json.load(f)
	
	for data in datas:
		if 'gal_num' in data and 'names' in data:
			df_Physical = df_Physical.append({'url':data['gal_num'],'Filename':data['names'][0]}, ignore_index=True)
		
	#print(df_Physical)
	
def mapping():
	global df_Spec
	global df_Physical
	
	for index, row in df_Physical.iterrows():
		res = df_Spec['url'][df_Spec['url']==row['url']].index.tolist()
		df_Spec.loc[res, 'htFilename'] = row['Filename']
		#print(res)
		
def rename(strBaseDir):
	global df_Spec
	global df_Physical
	
	for index, row in df_Spec.iterrows():
		if row[0] != 'x':
			strSheet = row['Sheet']
			iNo = int(row['No'])
			strNo = str(int(row['No'])).zfill(3)
			strTitle = (str(row['Title']).replace('?','')).replace(':', ' ')
			strArtist = str(row['Artist'])
			strHTFilename = row['htFilename']
			
			strTargetDir = os.path.join(strBaseDir, str(strSheet))
			if not os.path.exists(strTargetDir):
				os.makedirs(strTargetDir)

			strSourceFilename = os.path.join(strBaseDir, strHTFilename)
			strTargetFilename = os.path.join(strTargetDir, strNo+'.'+strTitle+'.mp3')
			
			print(strSourceFilename, strTargetFilename)
			shutil.copyfile(strSourceFilename, strTargetFilename)
			
			audio = eyed3.load(strTargetFilename)
			
			audio.tag.artist = strArtist
			audio.tag.album = strSheet
			audio.tag.title = strTitle
			audio.tag.track_num = (iNo, None)
	
			audio.tag.save(encoding='utf-8')
			

if __name__ == "__main__":

	# 宣告變數
	strXlsName = 'Music_from_Youtube.xlsx'
	lstSheets = ['TV']
	strHDT_Name = 'lists.hdt'
	strBaseDir = r'C:\Users\Tony\Downloads\hitomi_downloader_GUI'

	# 開始
	# prepare physical
	getPhysical(strHDT_Name)

	# prepare spec or read spec
	for strSheet in lstSheets:
		getUrls(strXlsName, str(strSheet))
		
	mapping()
	#print(df_Spec)
	
	rename(strBaseDir)
	
	
	





	
