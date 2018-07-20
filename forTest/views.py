# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from django.http import HttpResponse
from django.shortcuts import render
from django import template

from myproject.myapp.models import Document
from myproject.myapp.forms import DocumentForm

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
	xl.Workbooks.Close()
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
	
def detail(request):
		
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
	csv_reader = csv.reader(open('D:/DjangoWeb/myproject/media/DYNOTable.csv'))


	for line in csv_reader:
		# nAme, rEport, tYpe, cOde, pHase, dAte = line.split(', ')
		nAme = line[0]
		rEport = line[1].replace(" ","_")
		tYpe = line[2]
		cOde = line[3]
		pHase = line[4]
		dAte = line[5]
		mY = line[6]
		contributors.update([nAme])
		productCodes.update([cOde])
		loopCount += 1
		tableContent.append('<tr>\n')
		tableContent.append('<td>'+mY+'</td>\n')
		tableContent.append('<td>'+tYpe+'</td>\n')
		tableContent.append('<td>'+cOde+'</td>\n')
		tableContent.append('<td>'+pHase+'</td>\n')
		tableContent.append('''<td><a href="/media/documents/'''+rEport+'''.html" target="_blank">'''+rEport+'</a></td>\n')		
		tableContent.append('<td>'+dAte+'</td>\n')
		tableContent.append('<td>'+nAme+'</td>\n')		
		tableContent.append('</tr>\n')
			
	params['Detail_Table'] = ''.join(tableContent)
	params['Product_Codes'] = str(len(productCodes))
	params['Total_Report'] = str(loopCount)
	params['Contributors'] = str(len(contributors))
	
	# fp = open('templates/detail.html')
	# t = template.Template(fp.read())
	detail = []
	with open('templates/detail.html', 'r') as template:
		for line in template:
			detail.append(templateInstantiate(line, params))
	
	return HttpResponse(detail)

def upload(request):
	# Handle file upload
	
	if request.method == 'POST':
		docform = DocumentForm(request.POST, request.FILES)
		if docform.is_valid():
			newdoc = Document(docfile=request.FILES['docfile'])
	
			filename = docform.cleaned_data['docfile'].name.split('.')[0]
			fileformat = docform.cleaned_data['docfile'].name.split('.')[-1]
			name = docform.cleaned_data['your_name']
			
			my = filename.split('_')[0]			
			vp = filename.split('_')[1]
			code = filename.split('_')[2]
			phase = filename.split('_')[3]
			
			ftime = filename.split('_')[-1]
			fname = filename.replace(" ","_")
			localtime = timeNorm(ftime)
			# localtime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
			
			if fileformat == 'xlsx':
				newdoc.save()
				summaryTable = 'D:/DjangoWeb/myproject/media/DYNOTable.csv'
				fpath = 'D:/DjangoWeb/myproject/media/documents/' + fname
				convertExcel(fpath)
				csvfile = file(summaryTable,'ab+')
				writer = csv.writer(csvfile)
				writer.writerow([name,filename,vp,code,phase,localtime,my])
				csvfile.close()
			# else:
				# raise forms.ValidationError("gaga")

			# Redirect to the document list after POST
			
			
	else:
		docform = DocumentForm()  # A empty, unbound form
	# Load documents for the list page
	documents = Document.objects.all()

	# Render list page with the documents and the form
	return render(
		request,
		'upload.html',
		{'documents': documents, 'form': docform}
	)

def uploadtest(request):
	# Handle file upload
	
	if request.method == 'POST':
		docform = DocumentForm(request.POST, request.FILES)
		if docform.is_valid():
			newdoc = Document(docfile=request.FILES['docfile'])
	
			filename = docform.cleaned_data['docfile'].name.split('.')[0]
			fileformat = docform.cleaned_data['docfile'].name.split('.')[-1]
			name = docform.cleaned_data['your_name']
			
			my = filename.split('_')[0]			
			vp = filename.split('_')[1]
			code = filename.split('_')[2]
			phase = filename.split('_')[3]
			
			ftime = filename.split('_')[-1]
			fname = filename.replace(" ","_")
			localtime = timeNorm(ftime)
			# localtime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
			
			if fileformat == 'xlsx':
				pythoncom.CoInitialize() 
				newdoc.save()
				summaryTable = 'D:/DjangoWeb/myproject/media/DYNOTable.csv'
				fpath = 'D:/DjangoWeb/myproject/media/documents/' + fname
				csvfile = file(summaryTable,'ab+')
				writer = csv.writer(csvfile)
				writer.writerow([name,filename,vp,code,phase,localtime,my])
				csvfile.close()
				convertExcel(fpath)
			# else:
				# raise forms.ValidationError("gaga")

			# Redirect to the document list after POST
			
			
	else:
		docform = DocumentForm()  # A empty, unbound form
	# Load documents for the list page
	documents = Document.objects.all()

	# Render list page with the documents and the form
	return render(
		request,
		'upload.html',
		{'documents': documents, 'form': docform}
	)