import openpyxl
from utils import get_current_path,get_index_columns_config,get_currency,get_kwords_rowLimits_config,configToJson,get_tables_path
from utils import getSgvData,normalizeTable,pathsProyect
import os
import re
import json
import pandas as pd

paths=pathsProyect()
class scrapTablesExcel:
    def __init__(self,fileName,distributionType) -> None:
        self.XlsxPath=os.path.join(get_current_path(),"Cierres de Caja","formatoxlsx")
        self.indexColumns=get_index_columns_config()
        self.kwordsRowLimits=get_kwords_rowLimits_config()
        self.sh=openpyxl.load_workbook(os.path.join(get_current_path(),"Cierres de Caja","formatoxlsx",fileName)).worksheets[0]
        self.currency=get_currency(fileName)
        self.distributionType=distributionType
        self.sgvData=getSgvData(fileName)
        self.fileName=fileName
        self.codeCcaj=None
        self.currencyNeto=None
        self.get_recauda()
        self.gap=0
    def get_recauda(self):
        filename=self.fileName
        recaud=re.findall(r'(.*)_\d{5}_', filename)[0]
        self.recaud=recaud
        self.CodeCcaj=re.findall(r'_(\d{5})_', filename)[0]
        if self.currency.find("Bs")!=-1:
            self.currencyNeto="Bs"
        else:
            self.currencyNeto=self.currency
    def get_left_up_vertex_table(self,tableName,moneyType):
        sh=self.sh
        i=1
        columBill=self.indexColumns[self.distributionType]["Detalle Operaciones"][self.currency][moneyType][tableName]["valor"]
        kword=self.kwordsRowLimits[self.distributionType][self.currency][tableName]["superior"]
        while sh.cell(i,columBill).value !=kword:
            i=i+1
        return i
    def updateCurrency(self):
        sh=self.sh
        j=1
        i=6
        #move to the right until find a number
        while (sh.cell(i,j).value==None):
            j=j+1
        if j==2:
            self.currency="Bs00"
        elif j==4:
            self.currency="Bs"
        else:
            print("Error: no se pudo determinar el gap")
    
    def getBillstable(self):
        sh=self.sh
        filterBillsWords=["Cantidad",None]
        typeCurrency=self.currency
        typeFile=""
        typeDistribution=self.distributionType
        nameTable="billetes"
        columnsTableDict=self.indexColumns[typeDistribution]["Detalle Operaciones"][typeCurrency]["efectivo"][nameTable]
        columBill=columnsTableDict["valor"]
        columQantityBill=columnsTableDict["Cantidad"]
        ColumnAmountBill=columnsTableDict["subtotal"]
        columTotalBill=columnsTableDict["total"]
        billsTable=[]
        downLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["inferior"]
        moneyType="efectivo"
        i=self.get_left_up_vertex_table(nameTable,moneyType)
        while sh.cell(i,columTotalBill).value !=downLimitWord:
            billsDict={
                **self.sgvData,
                "Medio de pago":"Efectivo",
                "recaudadora":self.recaud,
                "moneda":self.currencyNeto,
                "ruta":self.fileName,
                "recaudacion":typeDistribution,
                "billValue":sh.cell(i,columBill).value,
                "billQuantity":sh.cell(i,columQantityBill).value,
                "Amount":sh.cell(i,ColumnAmountBill).value
            }
            if billsDict["billQuantity"] not in filterBillsWords:
                billsTable.append(billsDict)
            i=i+1
        #print(pd.DataFrame(billsTable))
        return billsTable

    def getCoinsTable(self):
        sh=self.sh
        nameTable="Monedas"
        filterCoinsWords=["Monedas",None]
        typeCurrency=self.currency
        typeDistribution=self.distributionType
        columnsTableDict=self.indexColumns[typeDistribution]["Detalle Operaciones"][typeCurrency]["efectivo"][nameTable]
        ColumCurrency=columnsTableDict["valor"]
        ColumnCurrencyQuantity=columnsTableDict["Cantidad"]
        ColumnCurrencyAmount=columnsTableDict["subtotal"]
        ColumnTotalCurrency=columnsTableDict["total"]
        i=self.get_left_up_vertex_table(nameTable,"efectivo")
        downLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["inferior"]
        coinsTable=[]
        while sh.cell(i,ColumnTotalCurrency).value !=downLimitWord:
            coinsDict={
                **self.sgvData,
                "Medio de pago":"Efectivo",
                "recaudadora":self.recaud,
                "moneda":self.currencyNeto,
                "ruta":self.fileName,
                "recaudacion":typeDistribution,
                "coinValue":sh.cell(i,ColumCurrency).value,
                "coinQuantity":sh.cell(i,ColumnCurrencyQuantity).value,
                "Amount":sh.cell(i,ColumnCurrencyAmount).value
            }
            if coinsDict["coinValue"] not in filterCoinsWords:
                coinsTable.append(coinsDict)
            i=i+1

        #print(pd.DataFrame(coinsTable))
        return coinsTable
    def getChecksTable(self):
        sh=self.sh
        nameTable="Cheques"
        moneyType="bancario"
        filterCheckWords=["Fecha",None,"Cheques"]
        typeCurrency=self.currency
        typeDistribution=self.distributionType
        columnsTableDict=self.indexColumns[typeDistribution]["Detalle Operaciones"][typeCurrency][moneyType][nameTable]
        ColumDate=columnsTableDict["fecha"]
        ColumCheckDocument=columnsTableDict["NroDocumento"]
        ColumnCheckBank=columnsTableDict["Banco"]
        ColumCheckAmount=columnsTableDict["subtotal"]
        ColumCheckTotal=columnsTableDict["total"]
        uplimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["superior"]
        downLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["inferior"]
        checkTable=[]
        i=1
        while sh.cell(i,ColumDate).value !=uplimitWord:
            i=i+1
        while sh.cell(i,ColumCheckTotal).value !=downLimitWord:
            checkDict={
                **self.sgvData,
                "Medio de pago":"Cheque",
                "recaudadora":self.recaud,
                "moneda":self.currencyNeto,
                "ruta":self.fileName,
                "recaudacion":typeDistribution,
                "Date":sh.cell(i,ColumDate).value,
                "DocumentNumber":sh.cell(i,ColumCheckDocument).value,
                "Bank":sh.cell(i,ColumnCheckBank).value,
                "Amount":sh.cell(i,ColumCheckAmount).value
            }
            if checkDict["Date"] not in filterCheckWords:
                checkTable.append(checkDict)
            i=i+1
        #print(pd.DataFrame(checkTable))
        return checkTable
    def getBankTransfersTable(self):
        sh=self.sh
        nameTable="transferencias"
        typeCurrency=self.currency
        moneyType="bancario"
        filterTransferWords=["Transferencias y/o Depósitos","Total Depósitos","Fecha",None]
        columnsTableDict=self.indexColumns[self.distributionType]["Detalle Operaciones"][typeCurrency][moneyType][nameTable]
        ColumBankTransfer=columnsTableDict["fecha"]
        ColumBankTransferDocument=columnsTableDict["NroDocumento"]
        ColumBankTransferBank=columnsTableDict["Banco"]
        ColumBankTransferAmount=columnsTableDict["subtotal"]
        ColumBankTransferTotal=columnsTableDict["total"]
        i=1
        upLimitWord=self.kwordsRowLimits[self.distributionType][typeCurrency][nameTable]["superior"]
        downLimitWord=self.kwordsRowLimits[self.distributionType][typeCurrency][nameTable]["inferior"]
        while sh.cell(i,ColumBankTransfer).value !=upLimitWord:
            i=i+1

        bankTransferTable=[]
        while sh.cell(i,ColumBankTransferTotal).value !=downLimitWord:
            bankTransferDict={
                **self.sgvData,
                "Medio de pago":"Transferencia Bancaria",
                "recaudadora":self.recaud,
                "moneda":self.currencyNeto,
                "ruta":self.fileName,
                "recaudacion":self.distributionType,
                "Date":sh.cell(i,ColumBankTransfer).value,
                "DocumentNumber":sh.cell(i,ColumBankTransferDocument).value,
                "Bank":sh.cell(i,ColumBankTransferBank).value,
                "Amount":sh.cell(i,ColumBankTransferAmount).value
            }
            if bankTransferDict["Date"] not in filterTransferWords:
                bankTransferTable.append(bankTransferDict)
            i=i+1
        #print(pd.DataFrame(bankTransferTable))
        return bankTransferTable
    def getSummaryTable(self):
        sh=self.sh
        columnsTableDict=self.indexColumns["distribuidora"]["Reporte Cobrador"]["ambos"]["reporte"]["reporte"]
        ColumCode=columnsTableDict["codigo"]
        ColumChecker=columnsTableDict["cobrador"]
        ColumRendDate=columnsTableDict["Fecha de Rend."]
        ColumReceiptDate=columnsTableDict["Fecha Recibo Emitido"]
        ColumReceiptNumber=columnsTableDict["NroRecibo"]
        ColumUsCash=columnsTableDict["EfectivoDolar"]
        ColumBsCash=columnsTableDict["EfectivoBs"]
        ColumUsCheck=columnsTableDict["ChequesDolar"]
        ColumBsCheck=columnsTableDict["ChequesBs"]
        ColumTotalUsCash=columnsTableDict["TotalEfectivoUs"]
        ColumTotalEqBsCash=columnsTableDict["TotalEfectivoEqBs"]
        ColumTotalBsCash=columnsTableDict["TotalEfectivoBs"]
        ColumTransferUs=columnsTableDict["TransfUs"]
        ColumTransferEqBs=columnsTableDict["TransfEqBs"]
        ColumTransferBs=columnsTableDict["TransfBs"]
        ColumTotalUs=columnsTableDict["CobradoUs"]
        ColumTotalEqBs=columnsTableDict["CobradoEqBs"]
        ColumTotalBs=columnsTableDict["CobradoBs"]
        ColumTotal=columnsTableDict["CobradoTotal"]
        
        upLimitWord=self.kwordsRowLimits["distribuidora"]["ambos"]["reporte"]["superior"]
        downLimitWord=self.kwordsRowLimits["distribuidora"]["ambos"]["reporte"]["inferior"]
        i=1
        summaryTable=[]
        filterkeyWords=[None,"Fecha de Rend."]
        while sh.cell(i,ColumRendDate).value !=upLimitWord:
            i=i+1

        while sh.cell(i,ColumRendDate).value !=downLimitWord:
            formatedDate=str(sh.cell(i,ColumReceiptDate).value).replace("/","")
            totalFormat=str(sh.cell(i,ColumTotal).value).replace(",","")
            checker=str(sh.cell(i,ColumChecker).value)
            summaryDict={
                "uniqKey":"",
                "ruta":self.fileName,
                "recaudadora":self.recaud,
                "Code":sh.cell(i,ColumCode).value,
                "Codigo":self.codeCcaj,
                "Checker":sh.cell(i,ColumChecker).value,
                "FechaRend":sh.cell(i,ColumRendDate).value,
                "FechaRecibo":sh.cell(i,ColumReceiptDate).value,
                "NroRecibo":sh.cell(i,ColumReceiptNumber).value,
                "UsCash":sh.cell(i,ColumUsCash).value,
                "BsCash":sh.cell(i,ColumBsCash).value,
                "UsCheck":sh.cell(i,ColumUsCheck).value,
                "BsCheck":sh.cell(i,ColumBsCheck).value,
                "TotalUsCash":sh.cell(i,ColumTotalUsCash).value,
                "TotalEqBsCash":sh.cell(i,ColumTotalEqBsCash).value,
                "TotalBsCash":sh.cell(i,ColumTotalBsCash).value,
                "TransferUs":sh.cell(i,ColumTransferUs).value,
                "TransferEqBs":sh.cell(i,ColumTransferEqBs).value,
                "TransferBs":sh.cell(i,ColumTransferBs).value,
                "TotalUs":sh.cell(i,ColumTotalUs).value,
                "TotalEqBs":sh.cell(i,ColumTotalEqBs).value,
                "TotalCCAJ":sh.cell(i,ColumTotalBs).value,
            }
            if summaryDict["FechaRend"] not in filterkeyWords:
                totalFormat="{:.2f}".format(float(totalFormat))
                rec=self.recaud
                summaryDict["uniqKey"]=rec+"_"+checker+"_"+formatedDate+"_"+totalFormat
                if float(totalFormat)>0:
                    summaryTable.append(summaryDict)
            i=i+1
        #print(pd.DataFrame(summaryTable))        
        return summaryTable

    def get_vouchers_table(self):
        sh=self.sh
        nameTable="voucher"
        typeCurrency=self.currency
        typeDistribution=self.distributionType
        columnstableDict=self.indexColumns["agencia"]["Detalle Operaciones"][typeCurrency]["eqEfectivo"][nameTable]
        columDate=columnstableDict["fecha"]
        columnNroRef=columnstableDict["Nro. Ref"]
        ColumNroClient=columnstableDict["Nro. CI."]
        ColumSubtotal=columnstableDict["subtotal"]
        ColumTotal=columnstableDict["total"]
        i=1
        voucherTable=[]
        kwordsFilter=[None,"Fecha","Total vouchers","Voucher de Tarjetas"]
        upLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["superior"]
        downLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["inferior"]
        while sh.cell(i,columDate).value !=upLimitWord:
            i=i+1
        while sh.cell(i,ColumTotal).value !=downLimitWord:
            voucherDict={
                **self.sgvData,
                "Medio de pago":"Voucher",
                "recaudadora":self.recaud,
                "moneda":self.currencyNeto,
                "ruta":self.fileName,
                "recaudacion":typeDistribution,
                "Date":sh.cell(i,columDate).value,
                "NroRef":sh.cell(i,columnNroRef).value,
                "NroClient":sh.cell(i,ColumNroClient).value,
                "Amount":sh.cell(i,ColumSubtotal).value,
            }
            if voucherDict["Date"] not in kwordsFilter:
                voucherTable.append(voucherDict)
            i=i+1
        #print(pd.DataFrame(voucherTable))
        return voucherTable
    def get_coupon_table(self):
        sh=self.sh
        nameTable="vales"
        typeCurrency=self.currency
        typeDistribution=self.distributionType
        columnstableDict=self.indexColumns[typeDistribution]["Detalle Operaciones"][typeCurrency]["eqEfectivo"][nameTable]
        columQuantity=columnstableDict["Cantidad"]
        columClient=columnstableDict["Cliente"]
        ColumSubtotal=columnstableDict["subtotal"]
        ColumTotal=columnstableDict["total"]

        i=1
        couponTable=[]
        keywordsFilter=[None,"Total vales"]
        upLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["superior"]
        downLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["inferior"]

        while sh.cell(i,columQuantity).value !=upLimitWord:
            i=i+1
        while sh.cell(i,ColumTotal).value !=downLimitWord:
            couponDict={
                **self.sgvData,
                "Medio de pago":"Vale",
                "recaudadora":self.recaud,
                "moneda":self.currencyNeto,
                "ruta":self.fileName,
                "recaudacion":self.distributionType,
                "Quantity":sh.cell(i,columQuantity).value,
                "Client":sh.cell(i,columClient).value,
                "Subtotal":sh.cell(i,ColumSubtotal).value,
            }
            i=i+1
        if couponDict["Quantity"] not in keywordsFilter:
            couponTable.append(couponDict)
            
        #print(pd.DataFrame(couponTable))
        return couponTable
    def get_qr_table(self):
        sh=self.sh
        nameTable="QR"
        typeDistribution=self.distributionType
        typeCurrency=self.currency
        columnstableDict=self.indexColumns["agencia"]["Detalle Operaciones"][typeCurrency]["eqEfectivo"][nameTable]
        columDate=columnstableDict["fecha"]
        columNroRef=columnstableDict["Nro. Ref"]
        ColumNroClient=columnstableDict["Nombre"]
        ColumSubtotal=columnstableDict["subtotal"]
        ColumTotal=columnstableDict["total"]

        i=1
        qrTable=[]
        keywordsFilter=[None,"Pagos QR","Fecha"]
        upLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["superior"]
        downLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["inferior"]
        while sh.cell(i,columDate).value !=upLimitWord:
            i=i+1
        
        while sh.cell(i,ColumTotal).value !=downLimitWord:
            qrDict={
                **self.sgvData,
                "Medio de pago":"QR",
                "recaudadora":self.recaud,
                "moneda":self.currencyNeto,
                "ruta":self.fileName,
                "recaudacion":self.distributionType,
                "Date":sh.cell(i,columDate).value,
                "NroRef":sh.cell(i,columNroRef).value,
                "NroClient":sh.cell(i,ColumNroClient).value,
                "Subtotal":sh.cell(i,ColumSubtotal).value,
            }
            if qrDict["Date"] not in keywordsFilter:
                qrTable.append(qrDict)
            i=i+1
        #print(pd.DataFrame(qrTable))
        return qrTable
    def get_diferences_table(self):
        typeCurrency=self.currency
        typeDistribution=self.distributionType
        sh=self.sh
        nameTable="diferencias"
        if typeCurrency=="Bs00":
            print(pd.DataFrame([]))
            return []
        columnstableDict=self.indexColumns[typeDistribution]["Detalle Operaciones"][typeCurrency]["contabilidad"][nameTable]
        columConcept=columnstableDict["Concepto"]
        columMotive=columnstableDict["Motivo"]
        ColumSubtotalUs=columnstableDict["subtotalDolar"]
        ColumSubtotalBs=columnstableDict["subtotalBs"]
        ColumTotalBs=columnstableDict["TotalBs"]
        ColumTotal=columnstableDict["TotalDiferencia"]

        upLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["superior"]
        downLimitWord=self.kwordsRowLimits[typeDistribution][typeCurrency][nameTable]["inferior"]
        
        i=1
        diferencesTable=[]
        keywordsFilter=[None,"Total Diferencias","Concepto"]
        while sh.cell(i,columConcept).value !=upLimitWord:
            i=i+1
        while sh.cell(i,ColumTotal).value !=downLimitWord:
            diferencesDict={
                **self.sgvData,
                "Medio de pago":"Diferencias",
                "recaudadora":self.recaud,
                "moneda":self.currencyNeto,
                "ruta":self.fileName,
                "Concept":sh.cell(i,columConcept).value,
                "Motive":sh.cell(i,columMotive).value,
                "SubtotalUs":sh.cell(i,ColumSubtotalUs).value,
                "SubtotalBs":sh.cell(i,ColumSubtotalBs).value,
                "TotalBs":sh.cell(i,ColumTotalBs).value,
            }
            
            if diferencesDict["Concept"] not in keywordsFilter:
                diferencesTable.append(diferencesDict)
            i=i+1
        #print(pd.DataFrame(diferencesTable))
        return diferencesTable
    def ventasEfectuadas(self):
        sh=self.sh
        ventasEfectuadas={}
        ventasEfectuadas["Ventas al Contado"]={}
        ventasEfectuadas["Ventas al Credito"]={}
        i=13
        j=1
        while sh.cell(i,j).value !='A':
            j=j+1
        ventasEfectuadas["Ventas al Contado"]['EFECTIVO']=sh.cell(i,j).value
        
        while sh.cell(i,j).value !='B':
            j=j+1
        ventasEfectuadas["Ventas al Contado"]['DOCUMENTOS']=sh.cell(i,j).value
        
        while sh.cell(i,j).value !='C=A+B':
            j=j+1
        ventasEfectuadas["Ventas al Contado"]['TOTAL']=sh.cell(i,j).value
        
        while sh.cell(i,j).value !='D':
            j=j+1   
        ventasEfectuadas["Ventas al Credito"]['PERSONAL IVSA']=''

        while sh.cell(i,j).value !='E':
            j=j+1
        ventasEfectuadas["Ventas al Credito"]['OTROS']=sh.cell(i,j).value
        
        while sh.cell(i,j).value !='F=D+E':
            j=j+1
        ventasEfectuadas["Ventas al Credito"]['TOTAL']=sh.cell(i,j).value
        
        while sh.cell(i,j).value !='G=C+F':
            j=j+1
        ventasEfectuadas['TOTALVENTAS']=''
        
        return ventasEfectuadas
def scrapXlsxFile(fileName):
    dicts=[]
    billTable=None
    checkTable=None
    bankTransferTable=None
    voucherTable=None
    cuoponTable=None
    qrTable=None
    coinsTable=None
    diferencesTable=None
    summaryTable=None
    if fileName.find("dist")!=-1:
        distributionType="distribuidora"
    elif fileName.find("ag")!=-1:
        distributionType="agencia"
    scrapyxlsx=scrapTablesExcel(fileName,distributionType)
    if distributionType=="distribuidora":
        if fileName.find("first")!=-1:
            summaryTable=scrapyxlsx.getSummaryTable()
            dictsTable={"summaryTable":summaryTable}    
        elif fileName.find("Us")!=-1:
            billTable=scrapyxlsx.getBillstable()
            checkTable=scrapyxlsx.getChecksTable()
            bankTransferTable=scrapyxlsx.getBankTransfersTable()
            dictsTable={"billTable":billTable,"checkTable":checkTable,"bankTransferTable":bankTransferTable}
        elif fileName.find("Bs")!=-1:
            billTable=scrapyxlsx.getBillstable()
            coinsTable=scrapyxlsx.getCoinsTable()
            checkTable=scrapyxlsx.getChecksTable()
            bankTransferTable=scrapyxlsx.getBankTransfersTable()
            dictsTable={"billTable":billTable,"coinsTable":coinsTable,"checkTable":checkTable,"bankTransferTable":bankTransferTable}
        else:
            print("Error en el nombre del archivo")
    elif distributionType=="agencia":
        if fileName.find("Us")!=-1:
            billTable=scrapyxlsx.getBillstable()
            voucherTable=scrapyxlsx.get_vouchers_table()
            qrTable=scrapyxlsx.get_qr_table()
            dictsTable={"billTable":billTable,"voucherTable":voucherTable,"qrTable":qrTable}
        elif fileName.find("Bs")!=-1:
            scrapyxlsx.updateCurrency()
            billTable=scrapyxlsx.getBillstable()
            coinsTable=scrapyxlsx.getCoinsTable()
            voucherTable=scrapyxlsx.get_vouchers_table()
            cuoponTable=scrapyxlsx.get_coupon_table()
            diferencesTable=scrapyxlsx.get_diferences_table()
            ventas=scrapyxlsx.ventasEfectuadas()
            dictsTable={"billTable":billTable,"coinsTable":coinsTable,"voucherTable":voucherTable,"cuoponTable":cuoponTable,"diferencesTable":diferencesTable}
        else:
            print("Error en el nombre del archivo")


    if billTable!=None:
        billst.extend(billTable)
    if checkTable!=None:
        checkstable.extend(checkTable)
    if bankTransferTable!=None:
        bankTransferstable.extend(bankTransferTable)
    if coinsTable!=None:
        coinssTable.extend(coinsTable)
    if voucherTable!=None:
        vouchersTable.extend(voucherTable)
    if qrTable!=None:
        qrsTable.extend(qrTable)
    if cuoponTable!=None:
        cuoponsTable.extend(cuoponTable)
    if diferencesTable!=None:
        diferencessTable.extend(diferencesTable)
    if summaryTable!=None:
        summariesTable.extend(summaryTable)
    dataReturn={"distributionType":distributionType,"data":dictsTable,"typeMoney":scrapyxlsx.currency}
    return dataReturn

def scrapCierresDeCaja():
    print("Procesando archivos de cierres de caja")
    #list of .xlsx files in the directory
    xlsxFilesList=[x for x in os.listdir(r"Cierres de Caja\formatoxlsx") if x.endswith(".xlsx")]
    global billst,checkstable,bankTransferstable,coinssTable,vouchersTable,qrsTable,cuoponsTable,diferencessTable,summariesTable
    billst=[]
    checkstable=[]
    bankTransferstable=[]
    vouchersTable=[]
    coinssTable=[]
    vouchersTable=[]
    qrsTable=[]
    cuoponsTable=[]
    diferencessTable=[]
    summariesTable=[]
    for xlsxFile in xlsxFilesList:
        print(xlsxFile)
        vd=scrapXlsxFile(xlsxFile)
    if len(xlsxFilesList)==0:
        print("No hay archivos xlsx en la carpeta Cierres de Caja")
        return
    df_bt=pd.DataFrame(billst)
    df_bt.to_csv(os.path.join(get_tables_path(),"billsTable.csv"),index=False,sep=";")

    df_checks=pd.DataFrame(checkstable)
    df_checks.to_csv(os.path.join(get_tables_path(),"checksTable.csv"),index=False,sep=";")

    df_bankTransfers=pd.DataFrame(bankTransferstable)
    df_bankTransfers.to_csv(os.path.join(get_tables_path(),"bankTransfersTable.csv"),index=False,sep=";")

    df_coins=pd.DataFrame(coinssTable)
    df_coins.to_csv(os.path.join(get_tables_path(),"coinsTable.csv"),index=False,sep=";")

    df_voucher=pd.DataFrame(vouchersTable)
    df_voucher.to_csv(os.path.join(get_tables_path(),"voucherTable.csv"),index=False,sep=";")

    df_qr=pd.DataFrame(qrsTable)
    df_qr.to_csv(os.path.join(get_tables_path(),"qrTable.csv"),index=False,sep=";")

    df_cuopon=pd.DataFrame(cuoponsTable)
    df_cuopon.to_csv(os.path.join(get_tables_path(),"cuoponTable.csv"),index=False,sep=";")

    df_summaries=pd.DataFrame(summariesTable)
    df_summaries.to_csv(os.path.join(get_tables_path(),"summariesTable.csv"),index=False,sep=";")

    normalizeTable()
    print("SCRAP CIERRES DE CAJA TERMINADO")

def scrapFiles():
    #configToJson()
    with open(r'src\target\CashClosingInfo.json',"r") as json_file:
        data = json.load(json_file)
    global billst,checkstable,bankTransferstable,coinssTable,vouchersTable,qrsTable,cuoponsTable,diferencessTable,summariesTable
    
    billst=[]
    checkstable=[]
    bankTransferstable=[]
    vouchersTable=[]
    coinssTable=[]
    vouchersTable=[]
    qrsTable=[]
    cuoponsTable=[]
    diferencessTable=[]
    summariesTable=[]
    
    for i,row in enumerate(data['data']):
        for j,path in enumerate(row["xlsFilesList"]):
            if path['descargado']=="OK":
                path=os.path.join(paths.dirCcaj,path['name']+".xlsx")
                vd=scrapXlsxFile(path)
                #data['data'][i]["xlsFilesList"][j]={}
                data['data'][i]["xlsFilesList"][j]["file"]=path
                data['data'][i]["xlsFilesList"][j]["distributionType"]=vd["distributionType"]
                data['data'][i]["xlsFilesList"][j]["moneyType"]=vd["typeMoney"]
                data['data'][i]["xlsFilesList"][j]["data"]=vd["data"]
                print(row['Código'])
            #except:
                #pass

    df_bt=pd.DataFrame(billst)
    df_bt.to_csv(os.path.join(get_tables_path(),"billsTable.csv"),index=False,sep=";")

    df_checks=pd.DataFrame(checkstable)
    df_checks.to_csv(os.path.join(get_tables_path(),"checksTable.csv"),index=False,sep=";")

    df_bankTransfers=pd.DataFrame(bankTransferstable)
    df_bankTransfers.to_csv(os.path.join(get_tables_path(),"bankTransfersTable.csv"),index=False,sep=";")

    df_coins=pd.DataFrame(coinssTable)
    df_coins.to_csv(os.path.join(get_tables_path(),"coinsTable.csv"),index=False,sep=";")

    df_voucher=pd.DataFrame(vouchersTable)
    df_voucher.to_csv(os.path.join(get_tables_path(),"voucherTable.csv"),index=False,sep=";")

    df_qr=pd.DataFrame(qrsTable)
    df_qr.to_csv(os.path.join(get_tables_path(),"qrTable.csv"),index=False,sep=";")

    df_cuopon=pd.DataFrame(cuoponsTable)
    df_cuopon.to_csv(os.path.join(get_tables_path(),"cuoponTable.csv"),index=False,sep=";")

    df_summaries=pd.DataFrame(summariesTable)
    df_summaries.to_csv(os.path.join(get_tables_path(),"summariesTable.csv"),index=False,sep=";")

    normalizeTable()
    print("SCRAP CIERRES DE CAJA TERMINADO")
    with open(r'src\target\FullExcelData.json', 'w') as outfile:
        json.dump(data, outfile,indent=4)
if __name__ == "__main__":
    scrapFiles()
    #scrapCierresDeCaja()
    #scrapCierresDeCobrador()
    #test_scrapDolarOperationsXls()