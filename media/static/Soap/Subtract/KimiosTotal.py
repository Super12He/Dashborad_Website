#!c:/Python27/python.exe
# -*- coding: utf-8 -*-
from suds.client import Client
from suds.plugin import *
import datetime
# import cgitb
# cgitb.enable()
from collections import UserDict
import logging
import json

logging.basicConfig(format='%(asctime)s %(message)s', level=logging.INFO, filename='D:\Documents\Scripts\Soap\KimiosReport.log')
logging.getLogger('suds.client').setLevel(logging.INFO)
logging.info('%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%')

Admin = Client('http://localhost:9999/kimios/services/AdministrationService?wsdl', cache=None)
Security = Client('http://localhost:9999/kimios/services/SecurityService?wsdl', cache=None)
SecurityXML = Client('http://localhost:9999/kimios/services/SecurityService?wsdl', cache=None, retxml=True)
Search = Client('http://localhost:9999/kimios/services/SearchService?wsdl', cache=None)
Document = Client('http://localhost:9999/kimios/services/DocumentVersionService?wsdl', cache=None)
DocumentAPI = Client('http://localhost:9999/kimios/services/DocumentService?wsdl', cache=None)
Folder = Client('http://localhost:9999/kimios/services/FolderService?wsdl', cache=None)
Studio = Client('http://localhost:9999/kimios/services/StudioService?wsdl', cache=None)
Token = Security.service.startSession('Xu Deyuan', 'PATAC', '123695')

fileName = 'DMS Report.txt'
fileNameTable = 'DMSTotal.txt'
# metaDefinition = {
        # 'Engine': {'Class': 503, 'Phase': 38, 'Type': 28},
        # 'Transmission': {'Class': 504, 'Phase': 38, 'Type': 29},
        # 'EE': {'Class': 5935, 'Phase': 38, 'Type': 41}
        # }
# documentType = {'Report': 501, 'Report: Engine': 503, 'Report: Transmission': 504, 'Report: EE': 5935}
dType = Studio.service.getDocumentTypes(Token).DocumentType
documentType = {}
for item in dType:
    documentType[item.name] = item.uid
# metaFeed[uid] = ['string', 'string']
mfType = Studio.service.getMetaFeeds(Token).MetaFeed
metaFeed= {}
for item in mfType:
    metaFeed[item.uid] = [str(a) for a in Studio.service.getMetaFeedValues(Token, item.uid).string]

responseTank = {}

class justDic():
    def __init__(self, anotherDic):
        self.owner = anotherDic.owner
        self.name = anotherDic.name
        self.extension = anotherDic.extension
        self.uid = anotherDic.uid
        self.folderUid = anotherDic.folderUid

    def __hash__(self):
        return hash(self.owner+self.name+self.extension+str(self.uid)+str(self.folderUid))

    def __eq__(self, other):
        return hash(self) == hash(other)

def reclass(ArrayOfDocument):
    reclassed = []
    for item in ArrayOfDocument:
        reclassed.append(justDic(item))
    return set(reclassed)

def advancedRetrieve(ArrayOfCriteria):
    sink = set([])
    for c in ArrayOfCriteria.Criteria:
        query = Search.factory.create('ns3:ArrayOfCriteria')
        query.Criteria = [c]
        response = set([])
        logging.info('query for %s' % c.fieldName+c.query)
        try:
            if c.fieldName+c.query not in responseTank:
                logging.info('new query insert into tank')
                responseTank[c.fieldName+c.query] = reclass(Search.service.advancedSearchDocuments(Token, query, 0, 10000000).rows.Document)
            response = responseTank[c.fieldName+c.query]
            logging.info('result obtained: %s' % len(response))
        except:
            responseTank[c.fieldName+c.query] = set([])
            logging.warning('empty query result')
            return []
        finally:
            if len(response) == 0:
                logging.warning('empty query result')
                return []
        # manual set Union/Intersect for AND/OR operators
        if len(sink) == 0:
            logging.info('sink initialized based on query %s' % c.fieldName+c.query)
            sink.update(response)
        else:
            if c.operator == 'AND':
                logging.info('sink intersection for query %s' % c.fieldName+c.query)
                logging.info('%s intersect %s' % (len(sink),len(response)))
                sink.intersection_update(response)
            elif c.operator == 'OR':
                logging.info('sink union for query %s' % c.fieldName+c.query)
                sink.update(response)
            else:
                logging.info('sink intersection for query %s' % c.fieldName+c.query)
                sink.intersection_update(response)
        if len(sink) == 0:
            logging.warning('empty sink returned')
            return []
    try:
        return sorted(list(sink),key=lambda item:item.owner)
    except:
        logging.warning('final empty sink returned')
        return []

def metaID(documentType, metaName):
#Given a meta value, return its meta uid
    ArrayofMeta = Document.service.getMetas(Token, documentType).Meta
    for item in ArrayofMeta:
        if item.metaFeedUid != -1:
            if metaName in metaFeed[item.metaFeedUid]:
                return item.uid

def queryGenerator(projectClass=None, ArrayOfString=None):
# projectClass should be listed in metaDefinitions, string projectPhase as in Alpha, beta, string projectType
    queryClass, queryType, queryPhase = (None, None, None)
    queryList = []
    if projectClass:
        duid = documentType[projectClass]
        queryClass = Search.factory.create('ns3:Criteria')
        queryClass.query = str(duid)
        queryClass.dateFacetGapType = None
        queryClass.dateFacetGapRange = None
        queryClass.level = 0
        queryClass.facetField = None
        queryClass.exclusiveFacet = False
        queryClass.metaId = None
        queryClass.metaType = None
        queryClass.fieldName = 'DocumentTypeUid'
        queryClass.filtersValues = None
        queryClass.rawQuery = False
        queryClass.faceted = False
        queryClass.operator = 'AND'
        queryList.append(queryClass)
    if ArrayOfString:
        for item in ArrayOfString:
            muid = metaID(duid, item)
            queryType = Search.factory.create('ns3:Criteria')
            queryType.query = item
            queryType.fieldName = 'MetaDataString_' + str(muid)
            queryType.position = 0
            queryType.level = 0
            queryType.exclusiveFacet = False
            queryType.metaId = muid
            queryType.metaType = 1
            queryType.filtersValues = None
            queryType.rawQuery = False
            queryType.faceted = False
            queryType.operator = 'AND'
            queryList.append(queryType)
    query = Search.factory.create('ns3:ArrayOfCriteria')
    query.Criteria = queryList
    return query

def reportTable(queryString, sink):
    with open(fileNameTable, 'a', encoding="utf-8") as fout:
        for item in sink:
            relatedDocument = DocumentAPI.service.getRelatedDocuments(Token, item.uid)
            if len(relatedDocument) != 0:
                try:
                    itemstr = [str(item.uid), str(item.owner).replace(' Archive', '').strip(), item.name + '.' + item.extension, relatedDocument.Document[0].name + '.' + relatedDocument.Document[0].extension]
                    itemstr.extend(queryString.replace('Report:', '').replace(' ','').split('&'))
                    fout.write('%s, %s, %s, %s, %s\n' % tuple(itemstr))
                except:
                    print(item.name)
                    continue
            else:
                itemstr = [str(item.uid), str(item.owner).replace(' Archive', '').strip(), item.name + '.' + item.extension, 'None']
                itemstr.extend(queryString.replace('Report:', '').replace(' ','').split('&'))
                fout.write('%s, %s, %s, %s, %s\n' % tuple(itemstr))

def reportGenerator(queryString, sink):
    with open(fileName, 'a', encoding="utf-8") as fout:
        fout.write('*** ' + ' ' + queryString.strip() + ' : ' + str(len(sink)) + ' ***\n')
        for item in sink:
            relatedDocument = DocumentAPI.service.getRelatedDocuments(Token, item.uid)
            if len(relatedDocument) != 0:
                try:
                    fout.write(str(item.owner).replace(' Archive', '').strip() + ': ' + item.name + '.' + item.extension + ': ' + relatedDocument.Document[0].name + '.' + relatedDocument.Document[0].extension + '\n')
                except:
                    print(item.name)
                    continue
            else:
                fout.write(str(item.owner).replace(' Archive', '').strip() + ': ' + item.name + '.' + item.extension + '\n')
                fout.write('~'*len(str(item.owner).replace(' Archive', '').strip() + ': ' + item.name + '.' + item.extension + '\n') + '\n')
        fout.write('*** ' + ' ' + queryString + ': END ***\n\n')
        fout.write('-'*120 + '\n')

# Demo
# projectClass = 'Report: Engine'
# queryString = ['NGC1.0T', 'Alpha']
# reportGenerator(advancedRetrieve(queryGenerator(projectClass, queryString)))

targets = ['Report']
with open(fileNameTable, 'w', encoding="utf-8") as fout:
    pass
with open(fileName, 'w', encoding="utf-8") as fout:
    fout.write('Report Generation Begin at %s \n' % datetime.datetime.now().isoformat())
for target in targets:
    queryString = target
    reportTable(queryString, advancedRetrieve(queryGenerator(target)))
with open(fileName, 'a', encoding="utf-8") as fout:
    fout.write('Report Generation End at %s \n' % datetime.datetime.now().isoformat())

