import sys
import os
from pathlib import Path
import pyexcel
import json
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

writeJson()