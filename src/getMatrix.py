#read FullExcelData.json
import json
import os
import sys
from utils import pathsProyect
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
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
    df_Agfinal=pd.DataFrame(matrixList)
    
    return df_Agfinal
def get_cobClientTable(ccobClients,cob,item):
    date_cob=cob['FechaRecibo_CcajConsol']
    excludeFields=["ruta_CcajCobClient","recaudadora_CcajCobClient"]
    cobCobTable=None
    if date_cob:
        date_cob=date_cob.replace("/","")
        amountCcob="{:.2f}".format(float(cob['TotalCcajConsol']))
        idCcob=f"{item['Recaudadora']}_{cob['Checker_CcajConsol']}_{date_cob}_{amountCcob}"
        filtro=filter(lambda x: x["ruta_CcajCobClient"] ==idCcob, ccobClients)
        cobCobTable=list(filtro)
        cobCobTable=[{k:v for k,v in d.items() if k not in excludeFields} for d in cobCobTable]
    if cobCobTable:
        pass
    else:
        if date_cob:
            cobCobTable=[{k:'info no encontrada' for k in ccobClients[0].keys() if k not in excludeFields} ]
        else:
            cobCobTable=[{k:'' for k in ccobClients[0].keys() if k not in excludeFields} ]
    return cobCobTable
def get_cobCobTable(ccobCob,cob,item):
    date_cob=cob['FechaRecibo_CcajConsol']
    excludeFields=["ruta_CcajCobCob","recaudadora_CcajCobCob"]
    cobCobTable=None
    if date_cob:
        date_cob=date_cob.replace("/","")
        amountCcob="{:.2f}".format(float(cob['TotalCcajConsol']))
        idCcob=f"{item['Recaudadora']}_{cob['Checker_CcajConsol']}_{date_cob}_{amountCcob}"
        filtro=filter(lambda x: x["ruta_CcajCobCob"] ==idCcob, ccobCob)
        cobCobTable=list(filtro)
        cobCobTable=[{k:v for k,v in d.items() if k not in excludeFields} for d in cobCobTable]
    if cobCobTable:
        pass
    else:
        if date_cob:
            cobCobTable=[{k:'info no encontrada' for k in ccobCob[0].keys() if k not in excludeFields} ]
        else:
            cobCobTable=[{k:'' for k in ccobCob[0].keys() if k not in excludeFields} ]
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
    totalCashOut=sum([float(str(x['TotalBsCash_CcajConsol']).replace(",","")) for x in ccobTable])

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
def get_CcajRecuentoTransf(CcobCobInfo,item):
    transferRecuentoTable=[]
    files=item["xlsFilesList"]
    for file in files:
        if file["descargado"]=="OK":
            if file["moneyType"]=="Bs" or file["moneyType"]=="Us":
                bankTransferTable_recuento=file['data']['bankTransferTable']
                checkTable_recuento=file['data']['checkTable']
    amounts_recuent=[x['AmountTransfer'] for x in bankTransferTable_recuento]
    if len(amounts_recuent)==0:
        transferRecuento={
                        "fechaTransf_CcajRecuen":"info no encontrada",
                        "NroTransf_CcajRecuen":"info no encontrada",
                        "BancoTransf_CcajRecuen":"info no encontrada",
                        "BsTransf_CcajRecuen":"info no encontrada",
                            }
        transferRecuentoTable.append(transferRecuento)
        return transferRecuentoTable
    
    for row in CcobCobInfo:
        if row['TransferBs_CcajCobCob'] in amounts_recuent:
            transferRecuento=bankTransferTable_recuento[amounts_recuent.index(row['TransferBs_CcajCobCob'])]
        else:
            if row['TransferBs_CcajCobCob']==0:
                transferRecuento={
                        "fechaTransf_CcajRecuen":0,
                        "NroTransf_CcajRecuen":0,
                        "BancoTransf_CcajRecuen":0,
                        "BsTransf_CcajRecuen":0,
                            }
            else:
                transferRecuento={
                        "fechaTransf_CcajRecuen":"info no encontrada",
                        "NroTransf_CcajRecuen":"info no encontrada",
                        "BancoTransf_CcajRecuen":"info no encontrada",
                        "BsTransf_CcajRecuen":"info no encontrada",
                            }
        transferRecuentoTable.append(transferRecuento)
    return transferRecuentoTable
def get_CcajRecuentoChecks(CcobCobInfo,item):
    files=item["xlsFilesList"]
    checksRecuentoTable=[]
    for file in files:
        if file["descargado"]=="OK":
            if file["moneyType"]=="Bs" or file["moneyType"]=="Us":
                bankTransferTable_recuento=file['data']['bankTransferTable']
                checkTable_recuento=file['data']['checkTable']
    amounts_recuent=[x['AmountCheck'] for x in checkTable_recuento]
    if len(amounts_recuent)==0:
        checkRecuento={
                    "fechaCheque_CcajRecuen":"",
                    "NroCheque_CcajRecuen":"",
                    "BancoCheque_CcajRecuen":"",
                    "BsCheque_CcajRecuen":"",
                        }
        checksRecuentoTable.append(checkRecuento)
        return checksRecuentoTable
    for row in CcobCobInfo:
        if row['CheckBs_CcajCobCob'] in amounts_recuent:
            checkRecuento=checkTable_recuento[amounts_recuent.index(row['TransferBs'])]
        else:
            checkRecuento={
                    "fechaCheque_CcajRecuen":"",
                    "NroCheque_CcajRecuen":"",
                    "BancoCheque_CcajRecuen":"",
                    "BsCheque_CcajRecuen":"",
                        }
        checksRecuentoTable.append(checkRecuento)
    return checksRecuentoTable
def get_CajRecuenTransfer2(CajaRecuentoTransferTable):
    cajaRecuentoTable=[]
    validInfo=False
    if "SapInfo" in CajaRecuentoTransferTable[0].keys():
        validInfo=True
    if CajaRecuentoTransferTable and validInfo:
        for row in CajaRecuentoTransferTable:
            cajaRecuentoDict={
            "fecha_CcajRecuenTransf":row['DateTransfer'],
            "BancoTransf_CcajRecuenTransf":row['DocumentNumberTransfer'],
            "BsTransf_CcajRecuenTransf":row['BankTransfer'],
            "UsTransf_CcajRecuenTransf":row['AmountTransfer'],
        }
        cajaRecuentoTable.append(cajaRecuentoDict)
    else:
        cajaRecuentoDict={
            "fecha_CcajRecuenTransf":"info no encontrada",
            "BancoTransf_CcajRecuenTransf":"info no encontrada",
            "BsTransf_CcajRecuenTransf":"info no encontrada",
            "UsTransf_CcajRecuenTransf":"info no encontrada"
        }
        cajaRecuentoTable.append(cajaRecuentoDict)
    return cajaRecuentoTable
def get_CcajRecuentoChecks2(CajaRecuentoChecksTable):
    checkSapTable=[]
    validInfo=False
    if "SapInfo" in CajaRecuentoChecksTable[0].keys():
        validInfo=True
    if CajaRecuentoChecksTable and validInfo:
        for row in CajaRecuentoChecksTable:
            cajaRecuentoDict={
            "fecha_CcajRecuenTransf":row['DateTransfer'],
            "BancoTransf_CcajRecuenTransf":row['DocumentNumberTransfer'],
            "BsTransf_CcajRecuenTransf":row['BankTransfer'],
            "UsTransf_CcajRecuenTransf":row['AmountTransfer'],
            }
            checkSapTable.append(cajaRecuentoDict)
    else:
        checkDict={
            "fecha_CcajRecuenCheck":"info no encontrada",
            "NroCheck_CcajRecuenCheck":"info no encontrada",
            "Banco_CcajRecuenCheck":"info no encontrada",
            "Bs_CcajRecuenCheck":"info no encontrada",
        }
        checkSapTable.append(checkDict)
    return checkSapTable
def get_SapInfoTransfer(CcajRecuentoTransfTable):
    transferTable=[]
    validInfo=False
    if 'SapInfo' in CcajRecuentoTransfTable[0].keys():
        validInfo=True
    if CcajRecuentoTransfTable and validInfo:
        for row in CcajRecuentoTransfTable:
            sapInfo=row['SapInfo']
            sapInfoDict={
                    "NroDocumentoTransfer_SapInfo":sapInfo['Nº doc.'],
                    "FechaDocuemntoTransfer_SapInfo":sapInfo['Fecha doc.'],
                    "ImporteBsTransfer_SapInfo":sapInfo['Importe en Ml'],
                    "TextoTransfer_SapInfo":sapInfo['Texto'],
                    "LibroMayorTransfer_SapInfo":sapInfo['Libro Mayor'],
                        }
            transferTable.append(sapInfoDict)
    else:
        sapInfoDict={
                "NroDocumentoTransfer_SapInfo":"info no encontrada",
                "FechaDocuemntoTransfer_SapInfo":"info no encontrada",
                "ImporteBsTransfer_SapInfo":"info no encontrada",
                "Texto_SapInfoTransfer":"info no encontrada",
                "LibroMayorTransfer_SapInfo":"info no encontrada",
                    }
        transferTable.append(sapInfoDict)
    return transferTable

def get_SapInfoChecks(CcajRecuentoChecksTable):
    checksTable=[]
    validInfo=False
    if 'SapInfo' in CcajRecuentoChecksTable[0].keys():
        validInfo=True
    if CcajRecuentoChecksTable and validInfo:
        for row in CcajRecuentoChecksTable:
            sapInfo=row['SapInfo']
            sapInfoDict={
                    "NroDocumentoCheck_SapInfo":sapInfo['Nº Doc.'],
                    "FechaDocuemntoCheck_SapInfo":sapInfo['Fecha Doc.'],
                    "ImporteBsCheck_SapInfo":sapInfo['Importe en Ml'],
                    "TextoCheck_SapInfo":sapInfo['Texto'],
                    "LibroMayorCheck_SapInfo":sapInfo['Libro Mayor'],
                        }
            checksTable.append(sapInfoDict)
    else:
        sapInfoDict={
                "NroDocumentoCheck_SapInfo":"info no encontrada",
                "FechaDocuemntoCheck_SapInfo":"info no encontrada",
                "ImporteBsCheck_SapInfo":"info no encontrada",
                "TextoCheck_SapInfo":"info no encontrada",
                "LibroMayorCheck_SapInfo":"info no encontrada",
                    }
        checksTable.append(sapInfoDict)
    return checksTable

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
                if cobRow['Code_CcajConsol']=="136":
                    pass
                sgvRow=get_SgvData(item)
                CcjaConsolidRow=cobRow
                sgvCashOutRow=get_SgvCashOut(cashOutInfo,ccobTable,item['Fecha'])               
                
                CcobCobTable=get_cobCobTable(CcobCob,cobRow,item)
                CcajRecuentoTransfTable=get_CcajRecuentoTransf(CcobCobTable,item)
                
                CcajRecuentoChecksTable=get_CcajRecuentoChecks(CcobCobTable,item)
                CobClientsTable=get_cobClientTable(CcobClients,cobRow,item)    
                
                sapInfoTableBankTransfer=get_SapInfoTransfer(CcajRecuentoTransfTable)
                CcajRecuentoTransfTable = get_CajRecuenTransfer2(CcajRecuentoTransfTable)

                sapInfoTableCheckTable=get_SapInfoChecks(CcajRecuentoChecksTable)
                CcajRecuentoChecksTable = get_CcajRecuentoChecks2(CcajRecuentoChecksTable)
                
                listOfTablesToConcat=[CcobCobTable,
                                      CcajRecuentoTransfTable,
                                      CcajRecuentoChecksTable,
                                      CobClientsTable,
                                      sapInfoTableBankTransfer,
                                      sapInfoTableCheckTable]
                concatTable=concat_dfs2(listOfTablesToConcat)

                for deepRow in concatTable:
                    finalDict={
                        **sgvRow,
                        **CcjaConsolidRow,
                        **sgvCashOutRow,
                        **deepRow,
                    }
                    matrixList.append(finalDict)
                    print(cobRow)
                    print("\n")
    df_Distfinal=pd.DataFrame(matrixList)
    df_Distfinal.to_csv(os.path.join(paths.folderProyect,'Tablas',"finalMatrixDist2.csv"),index=False,sep=";")            
    return df_Distfinal
    #print(pd.DataFrame(concatTable))            
    #concatTable.to_csv(os.path.join(paths.folderProyect,'Tablas',"finalMatrixDist.csv"),index=False,sep=";")

def makeFinalTemplate():
    df_Distfinal=getMatrixDist()
    df_Agfinal=getMatrixAg()
    pahtTemplate=os.path.join(paths.folderProyect,'Tablas',"plantillaRecaudadora.xlsx")
    namefile="finalTemplate.xlsx"
    pathOutTemplate=os.path.join(paths.folderProyect,'Ouputs',namefile)
    # Leer la plantilla Excel
    book = load_workbook(pahtTemplate)
    writer = pd.ExcelWriter(pathOutTemplate, engine='openpyxl')
    writer.book = book

    # Escribir el DataFrame en la hoja de la plantilla
    df_Distfinal.to_excel(writer, sheet_name='PT Distribuidora', index=False, startrow=6, header=False)
    df_Agfinal.to_excel(writer, sheet_name='PT Agencia ', index=False, startrow=6, header=False)
    # Guardar los cambios en el nuevo archivo Excel
    writer.close()
    pass
if __name__ == "__main__":
    makeFinalTemplate()
    pass