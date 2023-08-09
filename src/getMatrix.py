#read FullExcelData.json
import json
import os
import sys
from utils import pathsProyect,loginInfo
import pandas as pd
from datetime import datetime
from datetime import timedelta
from openpyxl import load_workbook
from utils import concat_dfs,concat_dfs2
paths=pathsProyect()

def get_SapInfoDocs(docAmount,recepit,sapInfo):
    sapDocs=[]
    sapInfo=[{**x, "Importe en Ml":round(float(x['Importe en Ml'].replace(".","").replace(",",".")),4)} for x in sapInfo]
    amountsSap=[x['Importe en Ml'] for x in sapInfo]
    founded=False
    if docAmount :
        if docAmount in amountsSap:
            #filter sapInfo with docAmount
            filtro=filter(lambda x: x["Importe en Ml"] ==docAmount, sapInfo)
            sapFiltered=list(filtro)
            if len(sapFiltered)>1:
                for row in sapFiltered:
                    if row['St']:
                        if recepit==row['St']:
                            sapInfoDict={
                                "NroDocumento_SapInfo":row['Nº doc.'],
                                "FechaDocuemnto_SapInfo":row['Fecha doc.'],
                                "ImporteBs_SapInfo":row['Importe en Ml'],
                                "Texto_SapInfo":row['Texto'],
                                "LibroMayor_SapInfo":row['Libro Mayor']
                            }
                            founded=True
                            break
                if founded==False:
                    sapInfoDict={
                        "NroDocumento_SapInfo":"info no encontrada",
                        "FechaDocuemnto_SapInfo":"info no encontrada",
                        "ImporteBs_SapInfo":"info no encontrada",
                        "Texto_SapInfo":"info no encontrada",
                        "LibroMayor_SapInfo":"info no encontrada",
                            }
            if len(sapFiltered)==1:

                index=amountsSap.index(sapFiltered[0]["Importe en Ml"])
                sapInfoDict={
                    "NroDocumento_SapInfo":sapInfo[index]['Nº doc.'],
                    "FechaDocuemnto_SapInfo":sapInfo[index]['Fecha doc.'],
                    "ImporteBs_SapInfo":sapInfo[index]['Importe en Ml'],
                    "Texto_SapInfo":sapInfo[index]['Texto'],
                    "LibroMayor_SapInfo":sapInfo[index]['Libro Mayor']
                }
            if len(sapFiltered)==0:
                sapInfoDict={
                    "NroDocumento_SapInfo":"info no encontrada",
                    "FechaDocuemnto_SapInfo":"info no encontrada",
                    "ImporteBs_SapInfo":"info no encontrada",
                    "Texto_SapInfo":"info no encontrada",
                    "LibroMayor_SapInfo":"info no encontrada",
                        }
        else:
            sapInfoDict={
                "NroDocumentoTransfer_SapInfo":"",
                "FechaDocuemntoTransfer_SapInfo":"",
                "ImporteBsTransfer_SapInfo":"",
                "Texto_SapInfoTransfer":"",
                "LibroMayorTransfer_SapInfo":"",
                    }
    else:
        sapInfoDict={
                "NroDocumentoTransfer_SapInfo":"",
                "FechaDocuemntoTransfer_SapInfo":"",
                "ImporteBsTransfer_SapInfo":"",
                "Texto_SapInfoTransfer":"",
                "LibroMayorTransfer_SapInfo":"",
                    }
        
    sapDocs.append(sapInfoDict)
    return sapDocs
def get_internTablesinfo(excelData):
    voucherTable=excelData['voucherTable']
    if not (excelData['voucherTable']):
        voucherTable=[
            {
            "DateVoucher":"info no encontrada",
            "NroRefVoucher":"info no encontrada",
            "NroClientVoucher":"info no encontrada",
            "AmountVoucher":"info no encontrada",
            }
        ]
    cuoponTable=excelData['cuoponTable']
    if not(excelData['cuoponTable']):
        cuoponTable=[
            {
            "QuantityVale":"info no encontrada",
            "ClientVale":"info no encontrada",
            "SubtotalVale":"info no encontrada",
            }
        ]
    qrTable=excelData['qrTable']
    if not(excelData['qrTable']):
        qrTable=[
            {
            "DateQr":"info no encontrada",
            "NroRefQr":"info no encontrada",
            "NroClientQr":"info no encontrada",
            "SubtotalQr":"info no encontrada",
            }
        ]
    diferencesTable=excelData['diferencesTable']
    if not(excelData['diferencesTable']):
        diferencesTable=[
            {
                "Concept":"info no encontrada",
                "Motive":"info no encontrada",
                "SubtotalUs":"info no encontrada",
                "SubtotalBs":"info no encontrada",
                "TotalBs":"info no encontrada",
            }
        ]
    return {"voucherTable":voucherTable,"valesTable":cuoponTable,
            "qrTable":qrTable,"diferencesTable":diferencesTable}

def getMatrixAg(loginData):
    print("Procesando archivo final...")
    sapLogin = loginData['SapLogin'] 
    fi = sapLogin['fechaInicio']
    fi = fi - timedelta(days=5)
    fi = fi.strftime("%d.%m.%Y")

    ff = sapLogin['fechaFin']
    ff = ff + timedelta(days=5)
    ff = ff.strftime("%d.%m.%Y")
    sapInfoCsvPath=os.path.join(paths.folderProyect,'SAPinfo',f"{fi} a {ff}.csv")
    df_sapInfo=pd.read_csv(sapInfoCsvPath,sep=";")
    sapInfoCsv=df_sapInfo.to_dict('records')
    jsonExcelDataPath=os.path.join(paths.folderProyect,'src','target','FullExcelData.json')
    #print(jsonExcelDataPath)
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
                    tablesInfo=get_internTablesinfo(excelData)
                    ventas=excelData['ventas']
                    docAmount=ventas['Ventas al Contado']['DOCUMENTOS']
                    receipt=item['Recibo']
                    otrosIngresos=excelData['otrosIngresos']

                    voucherTable=tablesInfo['voucherTable']
                    valesTable=tablesInfo['valesTable']
                    qrTable=tablesInfo['qrTable']
                    diferencesTable=tablesInfo['diferencesTable']

                    sapData=get_SapInfoDocs(docAmount,receipt,sapInfoCsv)
                    #list of tables 
                    matrixConcat=concat_dfs([voucherTable,valesTable,qrTable,diferencesTable,sapData])
                    
                    
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

                        "EFECTIVO":ventas['Ventas al Contado']['EFECTIVO'],
                        "DOCUMENTOS":ventas['Ventas al Contado']['DOCUMENTOS'],
                        "TOTAL":ventas['Ventas al Contado']['TOTAL'],
                        "PERSONAL IVSA":ventas['Ventas al Credito']['PERSONAL IVSA'],
                        "OTROS":ventas['Ventas al Credito']['OTROS'],
                        "TOTAL2":ventas['Ventas al Credito']['TOTAL'],
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
    df_Agfinal.to_csv(os.path.join(paths.folderProyect,'Tablas','Agfinal.csv'),index=False)
    #print(df_Agfinal)
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
        if amountCcob=="75328.91":
            pass
        idCcob=f"{item['Recaudadora']}_{cob['Checker_CcajConsol']}_{date_cob}_{amountCcob}"
        filtro=filter(lambda x: x["ruta_CcajCobCob"] ==idCcob, ccobCob)
        cobCobTable=list(filtro)
        cobCobTable=[{k:v for k,v in d.items() if k not in excludeFields} for d in cobCobTable]
    if cobCobTable:
        pass
    else:
        if date_cob:
            if date_cob.find("/")!=-1:
                cobCobTable=[{k:'info no encontrada' for k in ccobCob[0].keys() if k not in excludeFields} ]
            else:
                cobCobTable=[{k:v for k,v in d.items() if k not in excludeFields} for d in cobCobTable]
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
    totalCashOut=[float(str(x['TotalBsCash_CcajConsol']).replace(",","")) for x in ccobTable if x['FechaRend_CcajConsol']=="Saldo Final"]
    totalCashOut=round(sum(totalCashOut),4)
    sgvCashOut={
    "fecha de salida_SgvSaliE":"info no encontrada",
    "monto de salida_SgvSaliE":"info no encontrada",
    }
    cashOutsAmounts=[round(float(cashOut['Total Bs.'].replace(",","")),4) for cashOut in cashOutInfo]
    if totalCashOut in cashOutsAmounts:
        fechaSgvCcaj=cashOutInfo[cashOutsAmounts.index(totalCashOut)]['Fecha']
        #get object
        dateObjec1=datetime.strptime(fechaSgvCcaj, '%d/%m/%Y %H:%M:%S')
        fechaSgvCcaj=dateObjec1.strftime('%d/%m/%Y')
        sgvCashOut['fecha de salida_SgvSaliE']=cashOutInfo[cashOutsAmounts.index(totalCashOut)]['Fecha']
        sgvCashOut['monto de salida_SgvSaliE']=cashOutInfo[cashOutsAmounts.index(totalCashOut)]['Total Bs.']
        if fechaSgvCcaj==fechaSgvSaliE:
            pass
        else:
            pass
            #print("sin coincidencia de fechas")

    return sgvCashOut
def get_CcajRecuentoTransf(CcobCobInfo,item):
    transferRecuentoTable=[]
    files=item["xlsFilesList"]
    for file in files:
        if file["descargado"]=="OK":
            if "moneyType" in list(file.keys()):
                if file["moneyType"]=="Bs" or file["moneyType"]=="Us":
                    bankTransferTable_recuento=file['data']['bankTransferTable']
                    checkTable_recuento=file['data']['checkTable']
            else:
                bankTransferTable_recuento=[]
                checkTable_recuento=[]
    amounts_recuent=[x['AmountTransfer'] for x in bankTransferTable_recuento]
    if len(amounts_recuent)==0:
        transferRecuento={
                        "DateTransfer":"info no encontrada",
                        "DocumentNumberTransfer":"info no encontrada",
                        "BankTransfer":"info no encontrada",
                        "AmountTransfer":"info no encontrada",
                            }
        transferRecuentoTable.append(transferRecuento)
        return transferRecuentoTable
    
    for row in CcobCobInfo:
        if row['TransferBs_CcajCobCob']==21140:
            print("encontrado")
            pass
        
        if row['TransferBs_CcajCobCob'] in amounts_recuent:
            transferRecuento=bankTransferTable_recuento[amounts_recuent.index(row['TransferBs_CcajCobCob'])]

        else:
            if row['TransferBs_CcajCobCob']==0:
                transferRecuento={
                        "DateTransfer":0,
                        "DocumentNumberTransfer":0,
                        "BankTransfer":0,
                        "AmountTransfer":0,
                            }
            else:
                transferRecuento={
                        "DateTransfer":"info no encontrada",
                        "DocumentNumberTransfer":"info no encontrada",
                        "BankTransfer":"info no encontrada",
                        "AmountTransfer":"info no encontrada",
                            }
        transferRecuentoTable.append(transferRecuento)
    return transferRecuentoTable
def get_CcajRecuentoChecks(CcobCobInfo,item):
    files=item["xlsFilesList"]
    checksRecuentoTable=[]
    for file in files:
        if file["descargado"]=="OK":
            if "moneyType" in list(file.keys()):
                if file["moneyType"]=="Bs" or file["moneyType"]=="Us":
                    bankTransferTable_recuento=file['data']['bankTransferTable']
                    checkTable_recuento=file['data']['checkTable']
            else:
                bankTransferTable_recuento=[]
                checkTable_recuento=[]
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
        if "TransferBs" in row.keys():
            if row['CheckBs_CcajCobCob'] in amounts_recuent:
                checkRecuento=checkTable_recuento[amounts_recuent.index(row['TransferBs'])]
            else:
                checkRecuento={
                        "fechaCheque_CcajRecuen":"",
                        "NroCheque_CcajRecuen":"",
                        "BancoCheque_CcajRecuen":"",
                        "BsCheque_CcajRecuen":"",
                            }
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
    if CajaRecuentoTransferTable:
        for row in CajaRecuentoTransferTable:
            if "DateTransfer" in row.keys():
                cajaRecuentoDict={
                "fecha_CcajRecuenTransf":row['DateTransfer'],
                "BancoTransf_CcajRecuenTransf":row['DocumentNumberTransfer'],
                "BsTransf_CcajRecuenTransf":row['BankTransfer'],
                "UsTransf_CcajRecuenTransf":row['AmountTransfer'],
                }
            else:
                cajaRecuentoDict={
                "fecha_CcajRecuenTransf":"info no encontrada",
                "BancoTransf_CcajRecuenTransf":"info no encontrada",
                "BsTransf_CcajRecuenTransf":"info no encontrada",
                "UsTransf_CcajRecuenTransf":"info no encontrada"
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
    if CajaRecuentoChecksTable:
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
    emptySapInfo={
                "NroDocumentoTransfer_SapInfo":"info no encontrada",
                "FechaDocuemntoTransfer_SapInfo":"info no encontrada",
                "ImporteBsTransfer_SapInfo":"info no encontrada",
                "Texto_SapInfoTransfer":"info no encontrada",
                "LibroMayorTransfer_SapInfo":"info no encontrada",
                    }
    if CcajRecuentoTransfTable:
        for row in CcajRecuentoTransfTable:
            if "SapInfo" in list(row.keys()):
                sapInfo=row['SapInfo']
                if len(sapInfo)>0:
                    sapInfoDict={
                            "NroDocumentoTransfer_SapInfo":sapInfo['Nº doc.'],
                            "FechaDocuemntoTransfer_SapInfo":sapInfo['Fecha doc.'],
                            "ImporteBsTransfer_SapInfo":sapInfo['Importe en Ml'],
                            "TextoTransfer_SapInfo":sapInfo['Texto'],
                            "LibroMayorTransfer_SapInfo":sapInfo['Libro Mayor'],
                                }
                else:
                    sapInfoDict=emptySapInfo    
            else:
                sapInfoDict=emptySapInfo
            transferTable.append(sapInfoDict)
    else:
        sapInfoDict=emptySapInfo
        transferTable.append(sapInfoDict)
    return transferTable

def get_SapInfoChecks(CcajRecuentoChecksTable):
    checksTable=[]
    validInfo=False
    if CcajRecuentoChecksTable:
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
    with open(paths.jsonFinal,encoding="utf-8") as json_data:
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
                    if "moneyType" in list(file.keys()):
                        if file["moneyType"]=="Bs" or file["moneyType"]=="Us":
                            pass
                            #get the data of recuento
                        if file["moneyType"]=="other":
                            ccobTable=file["data"]['summaryTable']
                            #get the data of cierre de cobrador
                    else:
                        ccobTable=[]
                else:
                    ccobTable=[]
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
                                      CcajRecuentoChecksTable,
                                      CcajRecuentoTransfTable,
                                      CobClientsTable,
                                      sapInfoTableCheckTable,
                                      sapInfoTableBankTransfer
                                      ]
                concatTable=concat_dfs2(listOfTablesToConcat)

                for deepRow in concatTable:
                    finalDict={
                        **sgvRow,
                        **CcjaConsolidRow,
                        **sgvCashOutRow,
                        **deepRow,
                    }
                    matrixList.append(finalDict)
                    #print(cobRow)
                    #print("\n")
    df_Distfinal=pd.DataFrame(matrixList)
    #df_Distfinal.to_csv(os.path.join(paths.folderProyect,'Tablas',"finalMatrixDist2.csv"),index=False,sep=";")            
    return df_Distfinal
    #print(pd.DataFrame(concatTable))            
    #concatTable.to_csv(os.path.join(paths.folderProyect,'Tablas',"finalMatrixDist.csv"),index=False,sep=";")

def makeFinalTemplate(loginData):
    df_Distfinal=getMatrixDist()
    df_Agfinal=getMatrixAg(loginData)
    pahtTemplate=os.path.join(paths.folderProyect,'Tablas',"plantillaRecaudadora.xlsx")
    namefile="PlantillaFinal.xlsx"
    pathOutTemplate=os.path.join(paths.folderProyect,'Ouputs',namefile)
    # Leer la plantilla Excel
    wb =load_workbook(pahtTemplate)
    writer = pd.ExcelWriter(pathOutTemplate, engine='openpyxl')
    writer.book = wb
    # Escribir el DataFrame en la hoja de la plantilla
    df_Agfinal.to_excel(writer, sheet_name='PT Agencia ', index=False, startrow=6, header=False)
    df_Distfinal.to_excel(writer, sheet_name='PT Distribuidora', index=False, startrow=6, header=False)
    
    # Guardar los cambios en el nuevo archivo Excel
    writer.close()
    pass
if __name__ == "__main__":
    loginData=loginInfo()
    makeFinalTemplate(loginData)
    pass