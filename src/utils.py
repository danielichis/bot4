import sys
import os
from pathlib import Path
import pyexcel
import json
import pandas as pd
def get_current_path():
    config_name = 'myapp.cfg'
    # determine if application is a script file or frozen exe
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    application_path2 = Path(application_path)
    return application_path2.parent.absolute()

def delete_xlsFiles(folderPath):
    folderPathxls=os.path.join(folderPath,"descargas")
    for path in os.listdir(folderPathxls):
        # check if current path is a file
        if os.path.isfile(os.path.join(folderPathxls, path)):
            if path[-4:]==".xls" or path[-5:]==".xlsx":
                #if path=="auszug.txt" or path=="umsatz.txt":
                os.remove(os.path.join(folderPathxls, path))
                #print("txt file deleted")
    folderPathXlsx=os.path.join(folderPath,"descargasXlsx")
    for path in os.listdir(folderPathXlsx):
        # check if current path is a file
        if os.path.isfile(os.path.join(folderPathXlsx, path)):
            if path[-4:]==".xls" or path[-5:]==".xlsx":
                #if path=="auszug.txt" or path=="umsatz.txt":
                os.remove(os.path.join(folderPathXlsx, path))
                #print("txt file deleted")
def convert_xls(pathFolder):    
    donwloadFolder=os.path.join(pathFolder,"descargas")
    #C:\DanielBots\bot4\descargasXlsx
    filesInfolder=os.listdir(donwloadFolder)
    #filesFolder2=os.listdir(r"C:\DanielBots\bot4\descargasXlsx")
    e=""
    for file in filesInfolder:
        if file.endswith(".xls"):
            try:
                print(file)
                #name of file 
                xls=os.path.join(donwloadFolder,file)
                xlsx=os.path.join(pathFolder,"descargasXlsx",file.replace(".xls",".xlsx"))
                #xlsx=pathFolder+"\\descargasXlsx"+file.replace(".xls",".xlsx")
                #print(xls)
                #print(xlsx)
                pyexcel.save_book_as(file_name=xls, dest_file_name=xlsx)
            except Exception as e:
                print(e)
                os.remove(xlsx)
                #write_log(f"")
            #os.remove(xls)
    return e

def get_index_columns_config():
    tableJsonConfig=os.path.join(get_current_path(),"src","target","indexColumnsConfig.json")
    with open(tableJsonConfig) as json_file:
        data = json.load(json_file)
    return data
def get_kwords_rowLimits_config():
    tableJsonConfig=os.path.join(get_current_path(),"src","target","kwordsRowLimitsConfig.json")
    with open(tableJsonConfig) as json_file:
        data = json.load(json_file)
    return data
def get_currency(fileName):
    if fileName.find("Us")!=-1:
        typeCurrency="dólar"
    elif fileName.find("Bs")!=-1:
        typeCurrency="Bs"
    elif fileName.find("First")!=-1:
        typeCurrency="ambos"
    else:
        typeCurrency="other"
    return typeCurrency

def writeJson():
    with open(r'src\target\CashClosingInfo.json',"r") as json_file:
        data = json.load(json_file)
    for row in data['data']:
        row['NuevaData']={}
    with open(r'src\target\CashClosingInfo.json',"w") as json_file:
        json.dump(data,json_file,indent=4)
def configToJson():
    configxlsxPath=os.path.join(get_current_path(),"config.xlsx")
    indexColumnsPathJson=os.path.join(get_current_path(),"src","target","indexColumnsConfig.json")
    kwordsRowLimitsPathJson=os.path.join(get_current_path(),"src","target","kwordsRowLimitsConfig.json")

    dfC=pd.read_excel(configxlsxPath,sheet_name="columnas")
    #print(df)
    #conver the df into collection of dictionaries
    dataColumns=dfC.values.tolist()
    columnsDict = {}
    for d in dataColumns:
        if d[0] not in columnsDict:
            columnsDict[d[0]] = {}
        if d[1] not in columnsDict[d[0]]:
            columnsDict[d[0]][d[1]] = {}
        if d[2] not in columnsDict[d[0]][d[1]]:
            columnsDict[d[0]][d[1]][d[2]] = {}
        if d[3] not in columnsDict[d[0]][d[1]][d[2]]:
            columnsDict[d[0]][d[1]][d[2]][d[3]] = {}
        if d[4] not in columnsDict[d[0]][d[1]][d[2]][d[3]]:
            columnsDict[d[0]][d[1]][d[2]][d[3]][d[4]] = {}
        columnsDict[d[0]][d[1]][d[2]][d[3]][d[4]][d[5]] = d[6]
    with open(indexColumnsPathJson, 'w') as outfile:
        json.dump(columnsDict, outfile,indent=4)


    dfKwords=pd.read_excel(configxlsxPath,sheet_name="kwords")
    dataKeywords=dfKwords.values.tolist()
    kwordsDict = {}
    for d in dataKeywords:
        if d[0] not in kwordsDict:
            kwordsDict[d[0]] = {}
        if d[1] not in kwordsDict[d[0]]:
            kwordsDict[d[0]][d[1]] = {}
        if d[2] not in kwordsDict[d[0]][d[1]]:
            kwordsDict[d[0]][d[1]][d[2]] = {}
        kwordsDict[d[0]][d[1]][d[2]][d[3]] = d[4]

    with open(kwordsRowLimitsPathJson, 'w') as outfile:
        json.dump(kwordsDict, outfile,indent=4)
def findsums():
# Definimos las listas A y B
    A = [4, 2, 3, 8, 5]
    B = [13, 9]

    # Creamos una lista vacía para guardar la lista C
    C = []

    # Recorremos cada elemento de la lista B
    for b in B:
        # Creamos una lista vacía para guardar los elementos de A que suman b
        sublista = []
        # Inicializamos una variable para guardar la suma parcial
        suma = 0
        # Recorremos cada elemento de la lista A desde el principio hasta el final
        for a in A:
            # Si la suma parcial más el elemento actual es menor o igual que b
            if suma + a <= b:
                # Añadimos el elemento a la sublista y actualizamos la suma parcial
                sublista.append(a)
                suma += a
                # Si la suma parcial es igual que b
                if suma == b:
                    # Añadimos la sublista a la lista C y salimos del bucle interno
                    C.append(sublista)
                    break
            # Si la suma parcial más el elemento actual es mayor que b
            else:
                # Vaciamos la sublista y reiniciamos la suma parcial a cero
                sublista = []
                suma = 0
    return C
    # Imprimimos la lista C como resultado final

print(findsums())