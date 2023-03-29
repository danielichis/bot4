import sys
import os
from pathlib import Path
import pyexcel
import openpyxl
import json
import pandas as pd
import re
def get_current_path():
    config_name = 'myapp.cfg'
    # determine if application is a script file or frozen exe
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    application_path2 = Path(application_path)
    return application_path2.parent.absolute()
def get_tables_path():
    config_name = 'myapp.cfg'
    # determine if application is a script file or frozen exe
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    application_path2 = Path(application_path)
    return os.path.join(application_path2.parent.absolute(),"Tablas")


def remove_files(folderPath):
    for path in os.listdir(folderPath):
    # check if current path is a file
        if os.path.isfile(os.path.join(folderPath, path)):
            if path[-4:]==".xls" or path[-5:]==".xlsx":
                #if path=="auszug.txt" or path=="umsatz.txt":
                os.remove(os.path.join(folderPath, path))
                #print("txt file deleted")
def delete_xlsFiles(folderPath):
    remove_files(os.path.join(folderPath,"Cierres de Caja"))
    remove_files(os.path.join(folderPath,"Cierres de Caja","formatoxlsx"))
    remove_files(os.path.join(folderPath,"Cierres de Cobrador"))
    remove_files(os.path.join(folderPath,"Cierres de Cobrador","formatoxlsx"))
def convert_xls(pathFolder):    
    filesInfolder=os.listdir(pathFolder)
    e=""
    for file in filesInfolder:
        if file.endswith(".xls"):
            try:
                print(file)
                #name of file 
                xls=os.path.join(pathFolder,file)
                xlsx=os.path.join(pathFolder,"formatoxlsx",file.replace(".xls",".xlsx"))
                pyexcel.save_book_as(file_name=xls, dest_file_name=xlsx)
            except Exception as e:
                print(e)
                os.remove(xlsx)
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
def getSgvData(fileName):
    code=re.findall(r'(\d{5})_', fileName)[0]
    with open(r'src\target\FullExcelData.json',"r") as json_file:
        data = json.load(json_file)
    for d in data["data"]:
        if d['Código']==code:
            return d
def normalizeTable():
    df_bills=pd.read_csv(r'Tablas\billsTable.csv',sep=';')
    df_coins=pd.read_csv(r'Tablas\coinsTable.csv',sep=';')
    df_transfers=pd.read_csv(r'Tablas\banktransfersTable.csv',sep=';')
    df_vouchers=pd.read_csv(r'Tablas\voucherTable.csv',sep=';')

    df_all=pd.concat([df_bills,df_coins,df_transfers,df_vouchers],ignore_index=True)
    
    allData=df_all.to_dict('records')
    for d in allData:
        if d['Amount']=="-":
            d['Amount']=0
    df_all=pd.DataFrame(allData)
    df_all.to_csv(r'Tablas\allTable.csv',index=False,sep=';',header=True)

def loginInfo():
    wb=openpyxl.load_workbook(os.path.join(get_current_path(),"config.xlsx"))
    configDat={}
    ws=wb["login"]
    configDat['dates']={}
    configDat['dates']['dInit']=ws["B2"].value
    configDat['dates']['dEnd']=ws["B3"].value
    configDat['users']={}

    maxRow=ws.max_row
    for i in range(2,maxRow+1):
        if ws["G"+str(i)].value!=None and ws["F"+str(i)].value=="SI":
            configDat['users'][ws["G"+str(i)].value]={}
            configDat['users'][ws["G"+str(i)].value]['user']=ws["H"+str(i)].value
            configDat['users'][ws["G"+str(i)].value]['password']=ws["I"+str(i)].value

    configDat['flags']={}
    configDat['flags']['flow']=ws["B5"].value
    configDat['flags']['cumulative']=ws["B6"].value
    return configDat

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
if __name__ == '__main__':
    configData()