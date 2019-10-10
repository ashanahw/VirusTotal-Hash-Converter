import requests
import xlrd
import sys
import time
import json
import xlwt 
from xlwt import Workbook

#Open Input Excel Workbook
wb = xlrd.open_workbook(sys.argv[1]) 
sheet = wb.sheet_by_index(0)

#Extracting the first row
sheet.cell_value(0, 0)

#New Workbook for output
wbwrite = Workbook() 
sheet1 = wbwrite.add_sheet('Hashes') 
sheet1.write(0, 0, 'MD5') 
sheet1.write(0, 1, 'SHA-1') 
sheet1.write(0, 2, 'SHA-256') 



#VT API data
url = 'https://www.virustotal.com/vtapi/v2/file/report'
API_KEY = '<YOUR API KEY>'
HASH = ''

#Moving row by row down
for i in range(sheet.nrows): 
	
	try:
		HASH =(sheet.cell_value(i, 0))
		params = {'apikey': API_KEY, 'resource': HASH }
		response = requests.get(url, params=params)
		data = response.json()
		md5=data["md5"]	
		SHA1=data["sha1"]
		SHA256=data["sha256"]
        
		#Writing Data to new Excel sheet
		sheet1.write(i+1, 0, md5) 
		sheet1.write(i+1, 1, SHA1) 
		sheet1.write(i+1, 2, SHA256) 	
	except:
		continue
		
	print(i+1," of ",sheet.nrows," Completed")
#VirusTotal Public API allows only 4 requests per minute
	
	time.sleep(16)

wbwrite.save('HashConvertedOutput.xls')



#Written by Ashan Harindu
#z4rR0w