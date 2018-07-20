#!c:/Python27/python.exe
# -*- coding: utf-8 -*-
from suds.client import Client
import random, string
from suds.plugin import *
from datetime import datetime
import time
import copy

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

Admin = Client('http://localhost:9999/kimios/services/AdministrationService?wsdl', cache=None)
Security = Client('http://localhost:9999/kimios/services/SecurityService?wsdl', cache=None)
SecurityXML = Client('http://localhost:9999/kimios/services/SecurityService?wsdl', cache=None, retxml=True)
Search = Client('http://localhost:9999/kimios/services/SearchService?wsdl', cache=None)
Document = Client('http://localhost:9999/kimios/services/DocumentService?wsdl', cache=None)
Folder = Client('http://localhost:9999/kimios/services/FolderService?wsdl', cache=None)
Studio = Client('http://localhost:9999/kimios/services/WorkspaceService?wsdl', cache=None)
Token = Security.service.startSession('Xu Deyuan', 'PATAC', '123695')

# Select studios to process,
targetStudios = ['Performance CAE', 'Structural CAE', 'CAE Think Tank']
# targetStudios = ['PT Future']
# targetStudios = ['PT Future']
studios = Studio.service.getWorkspaces(Token)
selectedStudios = [studio for studio in studios[0] if studio.name in targetStudios]
# Select folders for iteration
foldersToProcess = []
for studio in selectedStudios:
    foldersToProcess.extend(Folder.service.getFolders(Token, studio.uid)[0])
# Process folders and items
counter = 0
for folder in foldersToProcess:
    subFolders, subDocuments = subItems(Token, folder.uid)
    for item in subDocuments:
        if item.length <= 1024*1024*200:
            counter += 1
            print('%s %s' % (counter,item.name))
            try:
                Document.service.checkinDocument(Token, item.uid)
            except:
                input("Restart Kimios, Press Enter to continue...")
                Token = Security.service.startSession('Xu Deyuan', 'PATAC', '123695')
                continue



