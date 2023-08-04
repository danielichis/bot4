import subprocess
import os
from pathlib import Path
from openpyxl import load_workbook
import win32com.client
from datetime import timedelta
from datetime import datetime
import json
import time
import pandas as pd
from utils import pathsProyect,loginInfo
changeTheDateInicio = None
changeTheDateFin = None
listOfAccounts = []

pths=pathsProyect()

def superTable(configData):
    print("Descargando informacion de SAP...")
    SAPinfoPath = os.path.join(pths.folderProyect, "SAPinfo")
    # currentPathParentFolder = Path(currenPath).parent
    sapLogin = configData['SapLogin']           
    listOfAccounts = configData['accounts']

    layout = sapLogin['layout']
    layout = layout.replace(" ","")

    if sapLogin['fechaInicio'] != None:
        changeTheDateInicio = True
    if sapLogin['fechaFin'] != None:
        changeTheDateFin = True

    proc = subprocess.Popen([sapLogin['SAPPath'], '-new-tab'])
    time.sleep(2)
    try: 
        sapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = sapGuiAuto.GetScriptingEngine
        connection = application.OpenConnection(sapLogin['environment'],True)
    except:
        try:
            print("ERROR REINICIANDO WIN32LCIENT...")
            proc.kill()
            time.sleep(2)
            proc = subprocess.Popen([sapLogin['SAPPath'], '-new-tab'])
            time.sleep(2)
            sapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = sapGuiAuto.GetScriptingEngine
            connection = application.OpenConnection(sapLogin['environment'],True)
        except:
            print("ERROR REINICIANDO WIN32LCIENT...")
            proc.kill()
            time.sleep(2)
            proc = subprocess.Popen([sapLogin['SAPPath'], '-new-tab'])
            time.sleep(2)
            sapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = sapGuiAuto.GetScriptingEngine
            connection = application.OpenConnection(sapLogin['environment'],True)


    #q=application.OpenConnection()
    session = connection.Children(0)

    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = sapLogin['user']
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = sapLogin['psw']
    session.findById("wnd[0]").sendVKey(0)

    session.EndTransaction()
    session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
    session.findById("wnd[0]").sendVKey(0)

    fi = sapLogin['fechaInicio']
    #fi = datetime.strptime(str(fi), "%d/%m/%Y")
    fi = fi - timedelta(days=5)
    fi = fi.strftime("%d.%m.%Y")

    ff = sapLogin['fechaFin']
    #ff = datetime.strptime(str(ff), "%d/%m/%Y")
    ff = ff + timedelta(days=5)
    ff = ff.strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = fi
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = ff
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 10
    session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()
    for i,j in enumerate(listOfAccounts):
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = j
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = i + 1

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = SAPinfoPath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{fi} a {ff}.txt"
    #session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    #session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 86
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    proc.kill()

def tableTransSap(configData):
    currentPath = pths.folderProyect
    SAPinfoPath = os.path.join(pths.folderProyect, "SAPinfo")
    # currentPathParentFolder = Path(currenPath).parent
    sapLogin = configData['SapLogin']
    fi=sapLogin['fechaInicio']
    ff=sapLogin['fechaFin']

    fi = sapLogin['fechaInicio']
    #fi = datetime.strptime(str(fi), "%d/%m/%Y")
    fi = fi - timedelta(days=5)
    fi = fi.strftime("%d.%m.%Y")

    ff = sapLogin['fechaFin']
    ff=ff+timedelta(days=5)
    #ff = datetime.strptime(str(ff), "%d/%m/%Y")
    ff = ff.strftime("%d.%m.%Y")
    fileName=f"{fi} a {ff}.txt"
    with open (os.path.join(currentPath,"Sapinfo",fileName)) as fp:
        lines=fp.readlines()
    tableLines=[]
    for i,line in enumerate(lines):
        #print(i)
        fields=line.split("\t")
        #print(f"-------{len(fields)}")
        if len(fields)==15 and fields[11]!='Texto':
            texto=fields[11]
            if texto=="":
                texto="-"

            lineDict={
                "St":fields[3],
                "Nº doc.":fields[5],
                "Fecha doc.":fields[6].strip(),
                "Fecha.contab":fields[7].strip(),
                "CT":fields[8],
                "Importe en Ml":fields[9].strip(),
                "ML":fields[10],
                "Texto":texto,
                "ImpteML2":fields[12].strip(),
                "ML2":fields[13],
                "Libro Mayor":fields[14].strip(),                
            }
            tableLines.append(lineDict)
            #insertDataToJsonAg(tableLines)
    df=pd.DataFrame(tableLines)
    df.to_csv(os.path.join(currentPath,"Sapinfo",f"{fi} a {ff}.csv"),index=False,sep=";",encoding="utf-8",header=True)

def searchInSapInfo(rowSgv,valuesSap):
    for rowSap in valuesSap	:
        amountSAp=rowSap['Importe en Ml'].replace(".","").replace(",",".")
        if float(amountSAp)==float(rowSgv['AmountTransfer']):
            if rowSap['Fecha doc.'].replace(".","/")==rowSgv['DateTransfer']:
                return rowSap
def searchInsapInfoFull(rowSgv,valuesSap):
    for rowSap in valuesSap	:
        amountSAp=rowSap['Importe en Ml'].replace(".","").replace(",",".")
        if float(amountSAp)==float(rowSgv['AmountTransfer']):
            d1=rowSap['Fecha doc.']
            #d1 to object date
            d1=datetime.strptime(d1, '%d.%m.%Y')
            d2=rowSgv['DateTransfer']
            d2=datetime.strptime(d2, '%d/%m/%Y')
            #get de  absolute diference in days
            diff=abs((d2-d1).days)
            if diff<=4:
                return rowSap
            
def insertDataToJsonAg(configData):
    currentPath = pths.folderProyect
    with open(os.path.join(currentPath,"src","target","FullExcelData.json"), "r") as read_file:
        data = json.load(read_file)
    datos=data['data']
    
    sapLogin = configData['SapLogin']
    fi=sapLogin['fechaInicio']
    #date to string format
    fi = fi - timedelta(days=5)
    fi=fi.strftime("%d.%m.%Y")
    ff=sapLogin['fechaFin']
    ff=ff+timedelta(days=5)
    ff=ff.strftime("%d.%m.%Y")
    currentPath = pths.folderProyect
    df_sap=pd.read_csv(os.path.join(currentPath,"Sapinfo",f"{fi} a {ff}.csv"),sep=";",encoding="utf-8")
    #df_sap['Importe en Ml']=df_sap['Importe en Ml'].astype(float)
    #nan values to null

    valuesSAp=df_sap.to_dict('records')

    
    for row in datos:
        if row['Acciones']=="distribuidora":
            files=row['xlsFilesList']
            for file in files:
                fieldsIn=["descargado","moneyType"]
                if set(fieldsIn).issubset(list(file.keys())):
                    if file["descargado"]=="OK" and file["moneyType"]=="Bs":
                        checktable=file['data']['checkTable']
                        bankTransferTable=file['data']['bankTransferTable']
                        if len(bankTransferTable)>0:
                            for i,rowt in enumerate(bankTransferTable):
                                sapinfo=searchInSapInfo(rowt,valuesSAp)
                                file['data']['bankTransferTable'][i]["SapInfo"]={}
                                if sapinfo:
                                    file['data']['bankTransferTable'][i]["SapInfo"]=sapinfo
                                else:
                                    #print(rowt['AmountTransfer'])
                                    #print("no se encontro con doble coincidencia en sap")
                                    sapinfo=searchInsapInfoFull(rowt,valuesSAp)
                                    if sapinfo:
                                        file['data']['bankTransferTable'][i]["SapInfo"]=sapinfo
                                    else:
                                        #print("no se encontro con coincidencia simple en sap")
                                        file['data']['bankTransferTable'][i]["SapInfo"]={}
                else:
                    print(f"ADVERTENGIA el archivo {file['name']} no tiene campo de moneda")

    with open(os.path.join(currentPath,"src","target","FinalDataToExcel.json"), "w",encoding="utf-8") as write_file:
        json.dump(data, write_file, indent=4)

    
    # data['tableLines']=tableTransSap()
    # with open(os.path.join(currentPath,"Sapinfo","data.json"), "w") as write_file:
    #     json.dump(data, write_file, indent=4)
    
if __name__ == '__main__':
    configData=loginInfo()
    #superTable(configData)
    print("holaaa")
    tableTransSap(configData)
    insertDataToJsonAg(configData)