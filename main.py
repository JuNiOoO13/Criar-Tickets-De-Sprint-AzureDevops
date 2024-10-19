import pandas as pd
import requests
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
    ask = str;
    itens = list;
    currentItem = int;
    maxListLen = int;
    allInputs = list;
    currentItens = list;
    allAllowedChars = '1234567890qwertyuiopasdfghjklçzxcvbnm';
    searchString = str;
    
    def __init__(self,ask,itens):
    
        self.ask = ask;
        self.itens = itens;
        self.currentItens = itens;
        self.maxListLen = len(itens);
        self.currentItem = 0;
        self.searchString = '';
        self.allInputs = {
            "up" : partial(self._movTo,-1),
            "down" : partial(self._movTo,1),
            "backspace" : partial(self.backSpaceFunction),
            "space" : partial(self.spaceFunction)
        }
        self._drawMenu();

    def startMenu(self):
        keyboard.on_press(self._on_key_event)
        keyboard.wait('enter');
        self.listIndex = self.itens.index(self.currentItens[self.currentItem])
        return self.listIndex;

    def _movTo(self,direction):
        if(not self.currentItem + direction > self.maxListLen - 1 and not self.currentItem + direction < 0):
            self.currentItem += direction;
    
    def spaceFunction(self):
        self.searchString += " ";
        self.searchElement();
        
    def backSpaceFunction(self):
        if not len(self.searchString) - 1 < 0:
            self.searchString = self.searchString[:len(self.searchString)-1]
        self.searchElement();
        self.currentItem = 0;

    def _drawMenu(self):
        system('cls');
        print(self.ask)
        for i,item in enumerate(self.currentItens):
            if(i == self.currentItem):
                print(" => " + str(item));
            else:
                print(" " + item);
        print('Pesquise Aqui: ' + self.searchString)

    def _on_key_event(self,event):
        if(event.name in self.allInputs.keys()):
            self.allInputs[event.name]();
            self._drawMenu();
        
        else:
            if(event.name in self.allAllowedChars):
                self.currentItem = 0;
                self.searchString += event.name;
                self.searchElement();

    def searchElement(self):
        self.currentItens = list(filter(lambda item: self.searchString.lower() in item.lower(),self.itens))
        self._drawMenu();


def getUserData():
    global sprint_path,produto,projeto
    sprintNames = []
    sprintPaths = [];

    avaliableSprints = getSprints();
    for item in avaliableSprints:
        sprintNames.append(item['name'])
        sprintPaths.append(item['path'])


    avaliableProjects = getFields('Projeto')
    avaliableProducts = getFields('Produto')

    sprint_path = getInfo('Deseja Escolher Qual Sprint: ',sprintNames,sprintPaths)
    produto = getInfo('Qual projeto: ',avaliableProducts['allowedValues']);
    projeto = getInfo('Qual Projeto: ',avaliableProjects['allowedValues'])

def getInfo(ask,data,dataToExtract=None):
    menu = Menu(ask,data)
    index = menu.startMenu();
    return data[index] if dataToExtract == None else dataToExtract[index];

def searchSpreadsheet():
    currentDir = getcwd();
    spreadSheetPath = join(currentDir, 'Planilhas')
    if(not exists(spreadSheetPath)):
        mkdir(spreadSheetPath);
    
    spreadSheets = listdir(spreadSheetPath);
    if(len(spreadSheets) == 0):
        print('Você não tem nenhuma planilha na pasta tente add e rodar novamente o Programa')
        exit()
    else:
        return spreadSheets;

def getSpreadsheetData(arqName):
    df = pd.read_excel(arqName, header=18, engine='openpyxl');
    objList = []
    for index, row in df.iterrows():
        if(not pd.isna(row['Título do ticket'])):
            obj = {}
            obj['name'] = row['Título do ticket']
            obj['description'] = row['Descrição']
            obj['effort'] = row['Horas']
            obj['acceptanceCriteria'] = row['Critério de aceitação']
            obj['taskType'] = 'Documentação' if pd.isna(row['Histórias de teste']) else 'Desenvolvimento'
            objList.append(obj);
            if(not pd.isna(row['QA'])):

                obj = {}
                obj['name'] = "QA - " + row['Título do ticket']
                obj['description'] =  row['Histórias de teste']
                obj['effort'] = row['QA']
                obj['acceptanceCriteria'] = row['Critério de aceitação']
                obj['taskType'] = 'QA'
                objList.append(obj);
    
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
            "value" : values['description'],
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

    response = requests.post(url, headers=headers, auth=auth, data=json.dumps(ticket))
    if response.status_code == 200 or response.status_code == 201:
        print("Ticket criado com sucesso!")
    else:
        print(f"Erro ao criar o ticket: {response.status_code}, {response.text}")

def getFields(fieldName):
    url = f'https://dev.azure.com/{organization}/{project}/_apis/wit/workitemtypes/Product Backlog Item/fields/{fieldName}?$expand=allowedvalues&api-version=7.2-preview.3'

    response = requests.get(url, auth=HTTPBasicAuth('', personal_access_token))
    if response.status_code == 200:
        field_data = response.json()
        if 'allowedValues' in field_data:
            return field_data;
        else:
            print('Esse campo não possui valores disponíveis.')
            exit();
    else:
        print(f'Erro: {response.status_code} - {response.text}')
        exit();

def getSprints():
    url = f'https://dev.azure.com/{organization}/{project}/{team}/_apis/work/teamsettings/iterations?api-version=7.0'

    response = requests.get(url, auth=HTTPBasicAuth('', personal_access_token))

    if response.status_code == 200:
        iterations = response.json()
        return iterations['value'];
    else:
        print(f'Erro: {response.status_code} - {response.text}')


def getProjects():
    url = f'https://dev.azure.com/{organization}/_apis/projects?api-version=7.0';
    response = requests.get(url, auth=HTTPBasicAuth('', personal_access_token))
    if response.status_code == 200:
        iterations = response.json()
        return iterations['value'];
    else:
        print(f'Erro: {response.status_code} - {response.text}')
    pass



def getJsonData():
    global team, personal_access_token, organization, project
    if(not exists('configs.json')):
        personal_access_token = input('Api-key => ');
        organization = input('organização =>  ');
    with open('configs.json', 'r') as file:
        pass


organization = 'agtechagro'
project = 'Atividades WEB'
team = 'Atividades WEB Team'
personal_access_token = ''
sprint_path = r'Atividades WEB\Comunicação PremoPlan - parte 1'
produto = 'PremoPlan'
projeto = 'PremoPlan 5.0'
dataFim = '18/10/2024'

url = f'https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/$Product%20Backlog%20Item?api-version=7.0'
headers = {
    'Content-Type': 'application/json-patch+json',
    'Accept': 'application/json'
}
auth = HTTPBasicAuth('', personal_access_token)



spreadSheets = searchSpreadsheet();

menu = Menu('Qual planinha:',spreadSheets)
index = menu.startMenu();
fullDir = join(getcwd(),'Planilhas',spreadSheets[index])
listaDados = getSpreadsheetData(fullDir);

getUserData()


for item in listaDados:
    createTicket(item);



