#read FullExcelData.json
import json
import os
import sys
from utils import pathsProyect
import pandas as pd
from utils import concat_dfs
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
                    excelData=file['data']
                    ventas=excelData['ventas']
                    otrosIngresos=excelData['otrosIngresos']

                    billTable=excelData['billTable']
                    coinsTable=excelData['coinsTable']
                    voucherTable=excelData['voucherTable']
                    cuoponTable=excelData['cuoponTable']
                    diferencesTable=excelData['diferencesTable']
                    #list of tables 
                    matrixConcat=concat_dfs([billTable,coinsTable,voucherTable,cuoponTable,diferencesTable])
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

if __name__ == "__main__":
    getMatrixAg()