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
	countfile = open('D:/DjangoWeb/myproject/media/count.dat','a+')
	dailyfile = open('D:/DjangoWeb/myproject/media/daily.dat','a+')
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

	detail = []
	with open('templates/index.html', 'r') as template:
		for line in template:
			detail.append(line)
	
	return HttpResponse(detail)


def upload(request):
	detail = []
	with open('templates/index.html', 'r') as template:
		for line in template:
			detail.append(line)
	
	return HttpResponse(detail)
