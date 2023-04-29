#read FullExcelData.json
import json
import os
import sys
from utils import pathsProyect
import pandas as pd
from datetime import datetime
from utils import concat_dfs,concat_dfs2
paths=pathsProyect()
def getMatrixAg():
    #read FullExcelData.json
    jsonExcelDataPath=os.path.join(paths.folderProyect,'src','target','FullExcelData.json')
    
    print(jsonExcelDataPath)
    with open(jsonExcelDataPath) as json_data:
        data = json.load(json_data)
    matrixList=[]
    for item in data['data']:
        if item['Acciones']=="agencia":
            files=item["xlsFilesList"]
            for file in files:
                if file["descargado"]=="OK" and file["moneyType"]=="Bs":
                    print(file['file'])
                    excelData=file['data']
                    ventas=excelData['ventas']
                    otrosIngresos=excelData['otrosIngresos']

                    billTable=excelData['billTable']
                    coinsTable=excelData['coinsTable']
                    voucherTable=excelData['voucherTable']
                    cuoponTable=excelData['cuoponTable']
                    diferencesTable=excelData['diferencesTable']
                    #list of tables 
                    matrixConcat=concat_dfs([voucherTable,cuoponTable,diferencesTable])
                    #concatenate all tables
                    for row in matrixConcat:
                        newFrame={
                        "Código":item['Código'],
                        "Recibo":item['Recibo'],
                        "Fecha":item['Fecha'],
                        "Fecha de Cierre":item['Fecha de Cierre'],
                        "Total (Bs.)":item['Total (Bs.)'],
                        "Fondo (Bs.)":item['Fondo (Bs.)'],
                        "Estado":item['Estado'],
                        "Cajero":item['Cajero'],
                        "Auditado Por":item['Auditado Por'],

                        "EFECTIVO":ventas['Ventas al Contado']['EFECTIVO'],
                        "DOCUMENTOS":ventas['Ventas al Contado']['DOCUMENTOS'],
                        "TOTAL":ventas['Ventas al Contado']['TOTAL'],
                        "PERSONAL IVSA":ventas['Ventas al Credito']['PERSONAL IVSA'],
                        "OTROS":ventas['Ventas al Credito']['OTROS'],
                        "TOTAL":ventas['Ventas al Credito']['TOTAL'],
                        "TOTAL VENTAS":ventas['TOTALVENTAS'],
                        "Total Efectivo en Bs.":otrosIngresos['totalEfectivoBs'],
                        "Total Recuento Moneda Extranjera en Bs":otrosIngresos['totalMEBs'],
                        "Fondo para cambios":otrosIngresos['fondosCambios'],
                        "ImporteDepositarBs":otrosIngresos['importeDepositarBs'],
                        "ImporteDepositarUs":otrosIngresos['importeDepositarUs'],
                       **row
                        }
                        matrixList.append(newFrame)

#convert to dataframe
    df_final=pd.DataFrame(matrixList)
    df_final.to_csv(os.path.join(paths.folderProyect,'Tablas',"finalMatrix.csv"),index=False,sep=";")

def get_cobClientTable(ccobClients,cob,item):
    date_cob=cob['FechaRecibo_CcajConsol']
    cobCobTable=None
    if date_cob:
        date_cob=date_cob.replace("/","")
        amountCcob="{:.2f}".format(float(cob['TotalCcajConsol']))
        idCcob=f"{item['Recaudadora']}_{cob['Checker_CcajConsol']}_{date_cob}_{amountCcob}"
        filtro=filter(lambda x: x["ruta"] ==idCcob, ccobClients)
        cobCobTable=list(filtro)
    if cobCobTable:
        pass
    else:
        if date_cob:
            cobCobTable=[{k:'info no encontrada' for k in ccobClients[0].keys()} ]
        else:
            cobCobTable=[{k:'' for k in ccobClients[0].keys()} ]
    return cobCobTable
def get_cobCobTable(ccobCob,cob,item):
    date_cob=cob['FechaRecibo_CcajConsol']
    cobCobTable=None
    if date_cob:
        date_cob=date_cob.replace("/","")
        amountCcob="{:.2f}".format(float(cob['TotalCcajConsol']))
        idCcob=f"{item['Recaudadora']}_{cob['Checker_CcajConsol']}_{date_cob}_{amountCcob}"
        filtro=filter(lambda x: x["ruta"] ==idCcob, ccobCob)
        cobCobTable=list(filtro)
    if cobCobTable:
        pass
    else:
        if date_cob:
            cobCobTable=[{k:'info no encontrada' for k in ccobCob[0].keys()} ]
        else:
            cobCobTable=[{k:'' for k in ccobCob[0].keys()} ]
    return cobCobTable
def get_SgvData(item):
    sgvData={
            "Código_SgvCcaj":item['Código'],
            "Recibo_SgvCcaj":item['Recibo'],
            "Fecha_SgvCcaj":item['Fecha'],
            "Fecha de Cierre_SgvCcaj":item['Fecha de Cierre'],
            "Total (Bs.)_SgvCcaj":item['Total (Bs.)'],
            "Estado_SgvCcaj":item['Estado'],
            "Cajero_SgvCcaj":item['Cajero'],
                }
    return sgvData

def get_SgvCashOut(cashOutInfo,ccobTable,fechaSgvSaliE):
    totalCashOut=sum([float(x['TotalBsCash_CcajConsol']) for x in ccobTable])

    sgvCashOut={
    "fecha de salida_SgvSaliE":"info no encontrada",
    "monto de salida_SgvSaliE":"info no encontrada",
    }
    for cashOut in cashOutInfo:
        if cashOut['Total Bs.']==totalCashOut:
            fechaSgvCcaj=cashOut['Fecha']
            #get object
            dateObjec1=datetime.strptime(fechaSgvCcaj, '%d/%m/%Y %H:%M:%S')
            fechaSgvCcaj=dateObjec1.strftime('%d/%m/%Y')
            sgvCashOut['fecha de salida_SgvSaliE']=cashOut['Fecha_CcajConsol']
            sgvCashOut['monto de salida_SgvSaliE']=cashOut['TotalBsCash_CcajConsol']
            if fechaSgvCcaj==fechaSgvSaliE:
                pass
            else:
                print("sin coincidencia de fechas")
            break

    return sgvCashOut
def get_CcajRecuento(CcobCobInfo,item):

    files=item["xlsFilesList"]
    for file in files:
        if file["descargado"]=="OK":
            if file["moneyType"]=="Bs" or file["moneyType"]=="Us":
                bankTransferTable_recuento=file['data']['bankTransferTable']
                checkTable_recuento=file['data']['checkTable']
    amounts_recuent=[x['AmountTransfer'] for x in bankTransferTable_recuento]
    for row in CcobCobInfo:
        if row['TransferBs'] in amounts_recuent:
            pass
    CcajRecuento={
            "fechaCheque_CcajRecuen":"",
            "NroCheque_CcajRecuen":"",
            "BancoCheque_CcajRecuen":"",
            "BsCheque_CcajRecuen":"",
            "fechaTransf_CcajRecuen":"",
            "NroTransf_CcajRecuen":"",
            "BancoTransf_CcajRecuen":"",
            "BsTransf_CcajRecuen":"",
                }
    return CcajRecuento
def get_sapInfo():
    pass
def getMatrixDist():
    with open(paths.jsonFinal) as json_data:
        data = json.load(json_data)
    matrixList=[]
    with open(paths.jsonCobBox) as json_data:
        CcobCob = json.load(json_data)
    with open(paths.jsonClientBox) as json_data:
        CcobClients = json.load(json_data)
    with open(paths.jsonCashOut) as json_data:
        cashOutInfo = json.load(json_data)

    for item in data['data']:
        if item['Acciones']=="distribuidora":
            files=item["xlsFilesList"]
            for file in files:
                if file["descargado"]=="OK":
                    if file["moneyType"]=="Bs" or file["moneyType"]=="Us":
                        pass
                        #get the data of recuento
                    if file["moneyType"]=="other":
                        ccobTable=file["data"]['summaryTable']
                        #get the data of cierre de cobrador
            for cobRow in ccobTable:
                sgvRow=get_SgvData(item)
                CcjaConsolidRow=cobRow
                sgvCashOutRow=get_SgvCashOut(cashOutInfo,ccobTable,item['Fecha'])               
                
                CcobCobTable=get_cobCobTable(CcobCob,cobRow,item)
                CcajRecuentoTable=get_CcajRecuento(CcobCobTable,item)
                CobClientsTable=get_cobClientTable(CcobClients,cobRow,item)    
                
                
                sapInfoTable=get_sapInfo()
                concatTable=concat_dfs2()
    df_final=pd.DataFrame(matrixList)
    df_final.to_csv(os.path.join(paths.folderProyect,'Tablas',"finalMatrixDist.csv"),index=False,sep=";")

if __name__ == "__main__":
    #getMatrixAg()
    getMatrixDist()