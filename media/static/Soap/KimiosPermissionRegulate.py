#!c:/Python27/python.exe
# -*- coding: utf-8 -*-
from suds.client import Client
import random, string
from suds.plugin import *
import logging
from datetime import datetime
import copy

logging.basicConfig(format='%(asctime)s %(message)s', level=logging.INFO, filename='D:\Documents\Scripts\Soap\KimiosPermissionRegulation.log')
logging.getLogger('suds.client').setLevel(logging.INFO)
logging.info('%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%')
forceOverride = False
logging.info('Force Override Flag: ' + str(forceOverride))
# Return all subfolders and documents under certain parent folder
# Depth First Method
def subItems(Token, uid):
    folders = Folder.service.getFolders(Token, uid)[0] if len(Folder.service.getFolders(Token, uid))!=0 else []
    documents = Document.service.getDocuments(Token, uid)[0] if len(Document.service.getDocuments(Token, uid))!=0 else []
    if len(folders) != 0:
        for folder in copy.copy(folders):
            subFolders, subDocuments = subItems(Token, folder.uid)
            folders.extend(subFolders)
            documents.extend(subDocuments)
    return folders, documents

# Built Read-Only DMSecurity Rule XML stream
def securityXML(ArrayOfDMEntitySecurity, isFolder):
    rawXML = """<?xml version="1.0" encoding="UTF-8"?>\n<security-rules>\n"""
    for item in ArrayOfDMEntitySecurity:
        if item.name == 'Admin@CAE': # Always writable for Admin
            rule = '<rule security-entity-type="'+ str(item.type) + '" security-entity-uid="' + item.name + '" security-entity-source="' + item.source + '" read="true" write="true" full="false" />\n'
        elif item.type == 1 and isFolder: # Writable while item is folder and rule is personal
            rule = '<rule security-entity-type="'+ str(item.type) + '" security-entity-uid="' + item.name + '" security-entity-source="' + item.source + '" read="true" write="true" full="false" />\n'
        else:
            rule = '<rule security-entity-type="'+ str(item.type) + '" security-entity-uid="' + item.name + '" security-entity-source="' + item.source + '" read="true" write="false" full="false" />\n'
        rawXML = rawXML + rule
    rawXML = rawXML + '</security-rules>'
    return rawXML

Admin = Client('http://localhost:9999/kimios/services/AdministrationService?wsdl', cache=None)
Security = Client('http://localhost:9999/kimios/services/SecurityService?wsdl', cache=None)
SecurityXML = Client('http://localhost:9999/kimios/services/SecurityService?wsdl', cache=None, retxml=True)
Search = Client('http://localhost:9999/kimios/services/SearchService?wsdl', cache=None)
Document = Client('http://localhost:9999/kimios/services/DocumentService?wsdl', cache=None)
Folder = Client('http://localhost:9999/kimios/services/FolderService?wsdl', cache=None)
Studio = Client('http://localhost:9999/kimios/services/WorkspaceService?wsdl', cache=None)
Token = Security.service.startSession('Xu Deyuan', 'PATAC', '123695')

# Generate archive users
currentUsers = Security.service.getUsers(Token,'PATAC')[0]
currentUsersUids = [item.uid for item in currentUsers]
#Admin.service.createGroup(Token, 'Archive', 'Archive', 'PATAC')
alphabet = string.ascii_letters + string.digits + string.punctuation
length = 6
randomGenerator = random.SystemRandom()
loopCount = 0
for user in currentUsers:
    if 'Archive' not in user.uid and user.uid+' Archive' not in currentUsersUids:
        logging.warning('New Shadow User Added: ' + user.uid + ' Archive')
        loopCount += 1
        pw = str().join(randomGenerator.choice(alphabet) for _ in range(length))
        Admin.service.createUser(Token, user.uid+' Archive', user.firstName, user.lastName, user.phoneNumber, user.mail, pw,'PATAC','0')
        Admin.service.addUserToGroup(Token, user.uid+' Archive', 'Archive', 'PATAC')
if loopCount == 0:
    logging.info('No Additional User Added')

# Select studios to process,
targetStudios = ['Performance CAE', 'Structural CAE']
# targetStudios = ['CAE Playground']
logging.info('The Selected Studio: ' + ', '.join(targetStudios))
studios = Studio.service.getWorkspaces(Token)
selectedStudios = [studio for studio in studios[0] if studio.name in targetStudios]
# Select folders for iteration
foldersToProcess = []
for studio in selectedStudios:
    foldersToProcess.extend(Folder.service.getFolders(Token, studio.uid)[0])
logging.info(str(len(foldersToProcess)) + ' Folders to Process')
# Process folders and items permission and ownership
today = datetime.today()
deltaLimit = 30
logging.info('Delta Limit Set to: ' + str(deltaLimit))
loopCount = 0
for folder in foldersToProcess:
    loopCount += 1;
    logging.info('Processing Folder: ' + folder.path)
    logging.info(str(len(foldersToProcess)-loopCount) + ' Left to Go')
    subFolders, subDocuments = subItems(Token, folder.uid)
    permissonsXmlStream = securityXML(Security.service.getDMEntitySecurities(Token, folder.uid)[0], True)
    for item in subFolders:
        dayDelta = (today - item.updateDate).days
        if dayDelta >= deltaLimit:
            if forceOverride or 'Archive' not in item.owner:
                logging.warning(item.path + ' Modifed: Change Owner & Update Permission')
                Admin.service.changeOwnership(Token, item.uid, item.owner.replace('Archive','').strip()+' Archive', item.ownerSource)
                Security.service.updateDMEntitySecurities(Token, item.uid, permissonsXmlStream, False)
    permissonsXmlStream = securityXML(Security.service.getDMEntitySecurities(Token, folder.uid)[0], False)
    for item in subDocuments:
        dayDelta = (today - item.versionUpdateDate).days
        if dayDelta >= deltaLimit:
            if forceOverride or 'Archive' not in item.owner:
                logging.warning(item.path + ' Modifed: Change Owner & Update Permission')
                Admin.service.changeOwnership(Token, item.uid, item.owner.replace('Archive','').strip()+' Archive', item.ownerSource)
                Security.service.updateDMEntitySecurities(Token, item.uid, permissonsXmlStream, False)
