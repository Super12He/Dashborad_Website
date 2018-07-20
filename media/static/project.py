#!D:/Python27/python.exe
# -*- coding: utf-8 -*-
import pg, re
def append2dic(Result, repSummary, contributors):
    for item in contributors:
        repSummary[item].append(0)
    for item in Result:
        repSummary[item['author']].pop()
        repSummary[item['author']].append(item['count'])

def templateInstantiate(line, params):
    placeHolders = re.findall('\{(.+?)\}', line)
    for item in placeHolders:
        if item in params:
            line = re.sub('\{'+item+'\}', params[item], line)
    return line

totalSQL = "Select meta_string_value.meta_string_value AS author, count(distinct document_version.document_id) AS count From document_version INNER JOIN meta_string_value ON document_version.id=meta_string_value.document_version_id Where (meta_string_value.meta_id IN (28, 29, 41)) Group By meta_string_value.meta_string_value"
totalDocSQL = "Select author, count(distinct document_id) From document_version Group By author"
yearSQL = "Select meta_string_value.meta_string_value AS author, count(distinct document_version.document_id) AS count From document_version INNER JOIN meta_string_value ON document_version.id=meta_string_value.document_version_id Where (document_version.modification_date >= CURRENT_TIMESTAMP - interval '1 year' and document_version.modification_date <= CURRENT_TIMESTAMP) AND (meta_string_value.meta_id IN (28, 29, 41)) Group By meta_string_value.meta_string_value"
quarterSQL = "Select meta_string_value.meta_string_value AS author, count(distinct document_version.document_id) AS count From document_version INNER JOIN meta_string_value ON document_version.id=meta_string_value.document_version_id Where (document_version.modification_date >= CURRENT_TIMESTAMP - interval '3 month' and document_version.modification_date <= CURRENT_TIMESTAMP) AND (meta_string_value.meta_id IN (28, 29, 41)) Group By meta_string_value.meta_string_value"
monthSQL = "Select meta_string_value.meta_string_value AS author, count(distinct document_version.document_id) AS count From document_version INNER JOIN meta_string_value ON document_version.id=meta_string_value.document_version_id Where (document_version.modification_date >= CURRENT_TIMESTAMP - interval '1 month' and document_version.modification_date <= CURRENT_TIMESTAMP) AND (meta_string_value.meta_id IN (28, 29, 41)) Group By meta_string_value.meta_string_value"

pgdbConn = pg.connect(dbname = 'kimios',host = '127.0.0.1', user = 'root', passwd = 'Xu@123')
totalResult = pgdbConn.query(totalSQL).dictresult()
totalDocResult = pgdbConn.query(totalDocSQL).dictresult()
yearResult = pgdbConn.query(yearSQL).dictresult()
quarterResult = pgdbConn.query(quarterSQL).dictresult()
monthResult = pgdbConn.query(monthSQL).dictresult()
pgdbConn.close()

totalDocNum = sum([i['count'] for i in totalDocResult])
totalRepNum = sum([i['count'] for i in totalResult])
contributorNum = len(totalDocResult)
params = {}
params['Total_Document'] = str(totalDocNum)
params['Total_Report'] = str(totalRepNum)
params['Contributors'] = str(contributorNum)
params['Link_Overview'] = 'Overview'
params['Link_Projects'] = 'Projects'
params['Link_Details'] = 'Details'
repSummary = {}
contributors = [i['author'] for i in totalResult]
for item in contributors:
    repSummary[item] = []
append2dic(monthResult, repSummary, contributors)
append2dic(quarterResult, repSummary, contributors)
append2dic(yearResult, repSummary, contributors)
append2dic(totalResult, repSummary, contributors)
loopCount = 0
tableContent = []
for item in contributors:
    loopCount += 1
    tableContent.append('<tr>\n')
    tableContent.append('<td>'+str(loopCount)+'</td>\n')
    tableContent.append('<td>'+item+'</td>\n')
    tableContent.append('<td>'+str(repSummary[item][0])+'</td>\n')
    tableContent.append('<td>'+str(repSummary[item][1])+'</td>\n')
    tableContent.append('<td>'+str(repSummary[item][2])+'</td>\n')
    tableContent.append('<td>'+str(repSummary[item][3])+'</td>\n')
    tableContent.append('</tr>\n')
params['Detail_Table'] = ''.join(tableContent)

with open('project.html', 'r') as template:
    for line in template:
        print templateInstantiate(line, params)
