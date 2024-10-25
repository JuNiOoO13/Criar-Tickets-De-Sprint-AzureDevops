import pandas as pd
import requests
import openpyxl
from requests.auth import HTTPBasicAuth
import json
import os
from os import getcwd, mkdir,listdir
from os.path import join, exists
import keyboard
from functools import partial
from os import system
from sys import exit


class Menu:
    ask = str
    itens = list
    currentItem = int
    maxListLen = int
    allInputs = list
    currentItens = list
    allAllowedChars = '1234567890qwertyuiopasdfghjklçzxcvbnm'
    searchString = str
    
    def __init__(self,ask,itens):
    
        self.ask = ask
        self.itens = itens
        self.currentItens = itens
        self.maxListLen = len(itens)
        self.currentItem = 0
        self.searchString = ''
        self.allInputs = {
            "up" : partial(self._movTo,-1),
            "down" : partial(self._movTo,1),
            "backspace" : partial(self.backSpaceFunction),
            "space" : partial(self.spaceFunction)
        }
        self._drawMenu()

    def startMenu(self):
        keyboard.on_press(self._on_key_event)
        keyboard.wait('enter')
        self.listIndex = self.itens.index(self.currentItens[self.currentItem])
        return self.listIndex

    def _movTo(self,direction):
        if(not self.currentItem + direction > self.maxListLen - 1 and not self.currentItem + direction < 0):
            self.currentItem += direction
    
    def spaceFunction(self):
        self.searchString += " "
        self.searchElement()
        
    def backSpaceFunction(self):
        if not len(self.searchString) - 1 < 0:
            self.searchString = self.searchString[:len(self.searchString)-1]
        self.searchElement()
        self.currentItem = 0

    def _drawMenu(self):
        system('cls')
        print(self.ask)
        for i,item in enumerate(self.currentItens):
            if(i == self.currentItem):
                print(" => " + str(item))
            else:
                print(" " + item)
        print('Pesquise Aqui: ' + self.searchString)

    def _on_key_event(self,event):
        if(event.name in self.allInputs.keys()):
            self.allInputs[event.name]()
            self._drawMenu()
        
        else:
            if(event.name in self.allAllowedChars):
                self.currentItem = 0
                self.searchString += event.name
                self.searchElement()

    def searchElement(self):
        self.currentItens = list(filter(lambda item: self.searchString.lower() in item.lower(),self.itens))
        self._drawMenu()

def getUserData():
    global sprint_path, produto, projeto, team,codeReview
    sprintNames = []
    sprintPaths = []

    avaliableTeams = getTeams()
    team = getInfo('Deseja Escolher qual Time: ',avaliableTeams);
    avaliableSprints = getSprints()
    for item in avaliableSprints:
        sprintNames.append(item['name'])
        sprintPaths.append(item['path'])


    avaliableProjects = getFields('Projeto')['allowedValues']
    avaliableProducts = getFields('Produto')['allowedValues']


    
    sprint_path = getInfo('Deseja Escolher Qual Sprint: ',sprintNames,sprintPaths)
    print(avaliableProducts)
    produto = getInfo('Qual produto: ',avaliableProducts + ['Nenhum'],avaliableProducts + [' '])
    projeto = getInfo('Qual Projeto: ',avaliableProjects + ['Nenhum'], avaliableProjects + [' '])
    codeReview = getInfo('Deseja Escolher Criar os Tickets de Code Reviews Separados? ', ['Sim','Não'],[True,False])
    
def getInfo(ask,data,dataToExtract=None):
    menu = Menu(ask,data)
    index = menu.startMenu()
    return data[index] if dataToExtract == None else dataToExtract[index]

def searchSpreadsheet():
    currentDir = getcwd()
    spreadSheetPath = join(currentDir, 'Planilhas')
    if(not exists(spreadSheetPath)):
        mkdir(spreadSheetPath)
    
    spreadSheets = listdir(spreadSheetPath)
    if(len(spreadSheets) == 0):
        print('Você não tem nenhuma planilha na pasta tente add e rodar novamente o Programa')
        exit()
    else:
        return spreadSheets

def getSpreadsheetData(arqName):
    df = pd.read_excel(arqName, header=18, engine='openpyxl')
    objList = []
    for index, row in df.iterrows():
        if(not pd.isna(row['Título do ticket'])):
            obj = {}
            obj['name'] = row['Título do ticket']
            obj['description'] = row['Descrição']
            obj['effort'] = row['Horas']
            obj['acceptanceCriteria'] = row['Critério de aceitação']
            obj['taskType'] = 'Documentação' if pd.isna(row['Histórias de teste']) else 'Desenvolvimento'
            objList.append(obj)
            if(not pd.isna(row['QA'])):

                obj = {}
                obj['name'] = "QA - " + row['Título do ticket']
                obj['description'] =  row['Histórias de teste']
                obj['effort'] = row['QA']
                obj['acceptanceCriteria'] = row['Critério de aceitação']
                obj['taskType'] = 'QA'
                objList.append(obj)
            if(codeReview):
                obj = {}
                obj['name'] = "CodeReview - " + row['Título do ticket']
                obj['description'] =  row['Histórias de teste']
                obj['effort'] = row['QA']
                obj['acceptanceCriteria'] = row['Critério de aceitação']
                obj['taskType'] = 'Code Review'
                objList.append(obj)
    
    return objList

def createTicket(values):
    ticket = [
        {
            "op" : "add",
            "path" : "/fields/System.Title",
            "value" : values['name']
        },
        {
            "op": "add",
            "path": "/fields/Microsoft.VSTS.Scheduling.Effort",
            "value": values['effort']
        },
        {
            "op" : "add",
            "path" : "/fields/System.Description",
            "value" : values['description']
        },
        {
            "op" : "add",
            "path" : "/fields/Microsoft.VSTS.Common.Priority",
            "value" : 2,
        },
        {
            "op" : "add",
            "path" : "/fields/Produto",
            "value" : produto,
        },
        {
            "op" : "add",
            "path" : "/fields/Projeto",
            "value" : projeto,
        },
        {
            "op" : "add",
            "path" : "/fields/Planejamento",
            "value" : "Planejado",
        },
        {
            "op" : "add",
            "path" : "/fields/Objetivo",
            "value" : "Inovação",
        },
        {
            "op" : "add",
            "path" : "/fields/Custom.Tipodatarefa",
            "value" : values['taskType'],
        },
        {
            "op": "add",
            "path": "/fields/System.IterationPath",
            "value": sprint_path
        }
     ]
    print(sprint_path)
    response = requests.post(url, headers=headers, auth=auth, data=json.dumps(ticket))
    exit()
    if response.status_code == 200 or response.status_code == 201:
        ticketData = response.json()
        print(f"Ticket {ticketData['id']} criado com sucesso!")
        return ticketData['id'], ticketData['taskType']
        

    else:
        print(f"Erro ao criar o ticket: {response.status_code}, {response.text}")
        return -1, 'erro'
    
def getFields(fieldName):
    url = f'https://dev.azure.com/{organization}/{project}/_apis/wit/workitemtypes/Product Backlog Item/fields/{fieldName}?$expand=allowedvalues&api-version=7.2-preview.3'

    response = requests.get(url, auth=HTTPBasicAuth('', personal_access_token))
    if response.status_code == 200:
        field_data = response.json()
        if 'allowedValues' in field_data:
            return field_data
        else:
            print('Esse campo não possui valores disponíveis.')
            exit()
    else:
        print(f'Erro: {response.status_code} - {response.text}')
        exit()

def getSprints():
    url = f'https://dev.azure.com/{organization}/{project}/{team}/_apis/work/teamsettings/iterations?api-version=7.0'

    response = requests.get(url, auth=HTTPBasicAuth('', personal_access_token))

    if response.status_code == 200:
        iterations = response.json()
        return iterations['value']
    else:
        print(f'Erro: {response.status_code} - {response.text}')

def getProjects():
    url = f'https://dev.azure.com/{organization}/_apis/projects?api-version=7.0'
    response = requests.get(url, auth=HTTPBasicAuth('', personal_access_token))
    if response.status_code == 200:



        iterations = response.json()
        return list(map(lambda x : x['name'],iterations['value']))
    else:
        print(f'Erro: {response.status_code} - {response.text}')

def getTeams():
    url = f'https://dev.azure.com/{organization}/_apis/teams?api-version=7.2-preview.3'
    print(url)
    response = requests.get(url, auth=HTTPBasicAuth('', personal_access_token))
    if response.status_code == 200:
        iterations = response.json()
        return list(map(lambda x : x['name'],iterations['value']))
    else:
        print(f'Erro: {response.status_code} - {response.text}')

def getJsonData():
    global team, personal_access_token, organization, project
    if(not exists('configs.json')):
        print('Primeira vez usando ferramenta necessario configurar: ')
        personal_access_token = input('Api-key => ')
        organization = input('organização =>  ')
        projectsList = getProjects()
        menu = Menu('Escolha seu Projeto: ',projectsList)
        choosedIndex = menu.startMenu()
        project = projectsList[choosedIndex]
        jsonData = {
            "organization" : organization,
            "api-key" : personal_access_token,
            "project" : project
        }

        with open('configs.json', 'w') as configFile:
            json.dump(jsonData, configFile,indent=4)

    else:
        with open('configs.json', 'r') as configFile:
            jsonData = json.load(configFile)
            personal_access_token = jsonData['api-key']
            organization = jsonData['organization']
            project = jsonData['project']

def setTicketsIdOnSpreadsheet(directory,itemList):
    wb = openpyxl.load_workbook(directory)
    ws = wb.active

    startLine = 20
    for index,item in enumerate(itemList):
        ws[f'B{startLine + index}'] = item

    wb.save(directory)



if(__name__ == '__main__'):
    organization = str
    project = str
    team = str
    personal_access_token = str
    sprint_path = str
    produto = str
    projeto = str
    dataFim = str
    codeReview = bool

    url = f'https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/$Product%20Backlog%20Item?api-version=7.0'
    headers = {
        'Content-Type': 'application/json-patch+json',
        'Accept': 'application/json'
    }
    auth = HTTPBasicAuth('', personal_access_token)


    getJsonData()
    getUserData()


    spreadSheets = searchSpreadsheet()

    menu = Menu('Qual planinha:',spreadSheets)
    index = menu.startMenu()
    fullDir = join(getcwd(),'Planilhas',spreadSheets[index])
    listaDados = getSpreadsheetData(fullDir)
    ticketsIds = []


    for item in listaDados:
        ticketId, ticketType = createTicket(item)
        if(not ticketId == -1):
            if(ticketType == 'QA' or ticketType == 'Code Review' ):
                    ticketsIds[len(ticketsIds) - 1] += f'/{ticketId}'
            else:
                ticketsIds.append(ticketId)

    setTicketsIdOnSpreadsheet(fullDir, ticketsIds)