# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from django.http import HttpResponse
from django.shortcuts import render
from django import template

from myproject.myapp.models import NVH
from myproject.myapp.forms import NVHForm

import re
import csv
import os
import time
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants,DispatchEx
import pythoncom


# Create your views here.	
def convertExcel (fname):
	pythoncom.CoInitialize() 
	yourExcelFile = fname + '.xlsx'
	newFileName = fname + '.html'
	xl = DispatchEx('Excel.Application')
	wb = xl.Workbooks.Open(yourExcelFile)
	wb.SaveAs(newFileName, constants.xlHtml)
	wb.Close(SaveChanges=0)
	xl.Quit()
	del xl

def timeNorm (time):
	year = time[0:4]
	month = time[4:6]
	day = time[6:8]
	Ntime = "%s-%s-%s"%(year,month,day)
	return Ntime
	
def templateInstantiate(line, params):
    placeHolders = re.findall('\{(.+?)\}', line)
    for item in placeHolders:
        if item in params:
            line = re.sub('\{'+item+'\}', params[item], line)
    return line

def getCount():
	countfile = open('D:/DjangoWeb/myproject/media/NVHcount.dat','a+')
	dailyfile = open('D:/DjangoWeb/myproject/media/NVHdaily.dat','a+')
	counttext = countfile.read()   
	dailytext = dailyfile.read()  
	dailycount = dailytext.split(" ")[0]
	dailydate = dailytext.split(" ")[1]
	localTime = time.strftime('%Y%m%d',time.localtime(time.time()))
	try:
		count = int(counttext)+1
	except:
		count = 1    

	if localTime != dailydate:
		dcount = 1	
	else:
		dcount = int(dailycount) + 1
	
	daily = str(dcount) + ' ' + localTime
	dailyfile.seek(0)
	dailyfile.truncate()
	dailyfile.write(str(daily))
	dailyfile.flush()
	dailyfile.close()
	
	countfile.seek(0)
	countfile.truncate()
	countfile.write(str(count))
	countfile.flush()
	countfile.close()	
	
	pv = 'You are the No.%s visitors today, %s vistors in history.' %(dcount,count)
	return pv

	
def detail(request):
		
	params = {}
	
	params['Link_Projects'] = 'Upload'
	params['Link_Overview'] = 'Overview'
	params['Link_Details'] = 'Details'

	loopCount = 0
	tableContent = []
	contributors = set([])
	productCodes = set([])
	# fIlename = 'CSS 50T Dyno Validation Daily Report-20171031.htm'
	# dAte = '2017/11/01'
	csv_reader = csv.reader(open('D:/DjangoWeb/myproject/media/NVHTable.csv'))


	for line in csv_reader:
		# nAme, rEport, tYpe, cOde, pHase, dAte = line.split(', ')
		# [code,phase,domain,filename,localtime,name]
		vP = line[0]
		cOde = line[1]
		pHase = line[2]
		dOmain = line[3]
		rEport = line[4].replace(" ","_")
		dAte = line[5]
		nAme = line[6]
		contributors.update([nAme])
		productCodes.update([cOde])
		loopCount += 1
		tableContent.append('<tr>\n')
		tableContent.append('<td>'+vP+'</td>\n')
		tableContent.append('<td>'+cOde+'</td>\n')
		tableContent.append('<td>'+pHase+'</td>\n')
		tableContent.append('<td>'+dOmain+'</td>\n')
		tableContent.append('''<td><a href="/media/NVH/'''+rEport+'''.html" target="_blank">'''+rEport+'</a></td>\n')		
		tableContent.append('<td>'+dAte+'</td>\n')
		tableContent.append('<td>'+nAme+'</td>\n')
		tableContent.append('</tr>\n')
			
	params['Detail_Table'] = ''.join(tableContent)
	params['Product_Codes'] = str(len(productCodes))
	params['Total_Report'] = str(loopCount)
	params['Contributors'] = str(len(contributors))
	pageView = getCount()
	params['Page_Views'] = pageView
	# fp = open('templates/detail.html')
	# t = template.Template(fp.read())
	detail = []
	with open('templates/NVHdetail.html', 'r') as template:
		for line in template:
			detail.append(templateInstantiate(line, params))
	
	return HttpResponse(detail)


def upload(request):
	# Handle file upload
	
	if request.method == 'POST':
		docform = NVHForm(request.POST, request.FILES)
		if docform.is_valid():
			newdoc = NVH(docfile=request.FILES['docfile'])
	
			filename = docform.cleaned_data['docfile'].name.split('.')[0]
			fileformat = docform.cleaned_data['docfile'].name.split('.')[-1]
			name = docform.cleaned_data['your_name']
			
			# my = filename.split('_')[0]			
			# vp = filename.split('_')[1]
			# code = filename.split('_')[2]
			# phase = filename.split('_')[3]
			
			vp = filename.split('_')[0]
			code = filename.split('_')[1]
			phase = filename.split('_')[2]
			domain = filename.split('_')[3]
			
			ftime = filename.split('_')[-1]
			fname = filename.replace(" ","_")
			localtime = timeNorm(ftime)
			# localtime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
			
			if fileformat == 'xlsx':
				newdoc.save()
				summaryTable = 'D:/DjangoWeb/myproject/media/NVHTable.csv'
				fpath = 'D:/DjangoWeb/myproject/media/NVH/' + fname
				convertExcel(fpath)
				csvfile = file(summaryTable,'ab+')
				writer = csv.writer(csvfile)
				# writer.writerow([name,filename,code,phase,localtime,my])
				writer.writerow([vp,code,phase,domain,filename,localtime,name])
				
				csvfile.close()
			# else:
				# raise forms.ValidationError("gaga")

			# Redirect to the document list after POST
			
			
	else:
		docform = NVHForm()  # A empty, unbound form
	# Load documents for the list page
	documents = NVH.objects.all()

	# Render list page with the documents and the form
	return render(
		request,
		'NVHupload.html',
		{'documents': documents, 'form': docform}
	)
