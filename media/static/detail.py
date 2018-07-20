#!D:/Python27/python.exe
# -*- coding: utf-8 -*-
import re
import csv
def templateInstantiate(line, params):
    placeHolders = re.findall('\{(.+?)\}', line)
    for item in placeHolders:
        if item in params:
            line = re.sub('\{'+item+'\}', params[item], line)
    return line

params = {}
params['Link_Overview'] = 'Overview'
params['Link_Projects'] = 'Upload'
params['Link_Details'] = 'Details'

loopCount = 0
tableContent = []
contributors = set([])
productCodes = set([])
# fIlename = 'CSS 50T Dyno Validation Daily Report-20171031.htm'
# dAte = '2017/11/01'
csv_reader = csv.reader(open('D:\CGI\Statistics\SummaryTable.csv'))


for line in csv_reader:
	# nAme, rEport, tYpe, cOde, pHase, dAte = line.split(', ')
	nAme = line[0]
	rEport = line[1]
	tYpe = line[2]
	cOde = line[3]
	pHase = line[4]
	dAte = line[5]
	contributors.update([nAme])
	productCodes.update([cOde])
	loopCount += 1
	tableContent.append('<tr>\n')
	tableContent.append('<td>'+nAme+'</td>\n')
	tableContent.append('<td>'+tYpe+'</td>\n')
	tableContent.append('<td>'+cOde+'</td>\n')
	tableContent.append('<td>'+pHase+'</td>\n')
	tableContent.append('''<td><a href="/media/static/uploads/'''+rEport+'''.html">'''+rEport+'</a></td>\n')
	tableContent.append('<td>'+dAte+'</td>\n')
	tableContent.append('</tr>\n')
		
params['Detail_Table'] = ''.join(tableContent)
params['Product_Codes'] = str(len(productCodes))
params['Total_Report'] = str(loopCount)
params['Contributors'] = str(len(contributors))

detail = []

with open('D:/CGI/Statistics/detail.html', 'r') as template:
    for line in template:
        detail.append(templateInstantiate(line, params))

print detail
