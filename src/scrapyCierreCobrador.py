import openpyxl
from utils import get_current_path,get_index_columns_config,get_currency,get_kwords_rowLimits_config,configToJson
import os
import re
import json
import pandas as pd

class scraperCierreCobrador():
    def __init__(self,fileName) -> None:
        self.XlsxPath=os.path.join(get_current_path(),"Cierres de Cobrador","formatoxlsx",fileName)
        self.indexColumns=get_index_columns_config()
        self.kwordsRowLimits=get_kwords_rowLimits_config()
        self.sh=openpyxl.load_workbook(self.XlsxPath).worksheets[0]
        #self.currency=get_currency(fileName)
        #self.distributionType=distributionType
        #self.gap=0
    def ClientToCollectorTable(self):
        tableColumns=self.indexColumns['distribuidora']['Cierre de cobrador']
        tableKwords=self.kwordsRowLimits['distribuidora']['ambos']['recibo de cobranza']
        upperLimit=tableKwords['superior']
        lowerLimit=tableKwords['inferior']
        leftColumn=tableColumns['ambos']['recibo de caja']['recibo de caja']['Nro APP']

        appNumber=leftColumn
        recepitDate=tableColumns['ambos']['recibo de caja']['recibo de caja']['Fecha Recibo']
        clientCode=tableColumns['ambos']['recibo de caja']['recibo de caja']['Cod Cliente']
        clientName=tableColumns['ambos']['recibo de caja']['recibo de caja']['Nombre cliente']
        bsAmount=tableColumns['Bs']['efectivo']['recibo de caja']['Bs']
        UsAmount=tableColumns['Us']['efectivo']['recibo de caja']['Us']

        checkDate=tableColumns['ambos']['Cheque']['recibo de caja']['fechaCheque']
        checkNumber=tableColumns['ambos']['Cheque']['recibo de caja']['Nro Cheque']
        checkBank=tableColumns['ambos']['Cheque']['recibo de caja']['Banco Cheque']
        bsCheck=tableColumns['Bs']['Cheque']['recibo de caja']['Bs']
        usCheck=tableColumns['Us']['Cheque']['recibo de caja']['Us']

        dateTransfer=tableColumns['ambos']['transferencia']['recibo de caja']['Fecha Transferencia']
        bankTransfer=tableColumns['ambos']['transferencia']['recibo de caja']['Banco Transferencia']
        bsTransfer=tableColumns['Bs']['transferencia']['recibo de caja']['Bs']
        usTransfer=tableColumns['Us']['transferencia']['recibo de caja']['Us']

        subtotalBs=tableColumns['Bs']['subtotal']['recibo de caja']['Bs']
        subtotalUs=tableColumns['Us']['subtotal']['recibo de caja']['Us']
        subtotalEqBs=tableColumns['Bs']['subtotal']['recibo de caja']['EqBs']
        totalBs=tableColumns['Bs']['total']['recibo de caja']['Bs']
        reciboDeCajaTable=[]
        filtersKwords=['Nro APP','Nº de APP']
        i=1
        while self.sh.cell(row=i,column=leftColumn).value!=upperLimit:
            i+=1

        while self.sh.cell(row=i,column=clientCode).value!=lowerLimit:
            ditTable={
                'Nro APP':self.sh.cell(row=i,column=appNumber).value,
                'Fecha Recibo':self.sh.cell(row=i,column=recepitDate).value,
                'Cod Cliente':self.sh.cell(row=i,column=clientCode).value,
                'Nombre cliente':self.sh.cell(row=i,column=clientName).value,
                'CashBs':self.sh.cell(row=i,column=bsAmount).value,
                'CashUs':self.sh.cell(row=i,column=UsAmount).value,
                'CheckDate':self.sh.cell(row=i,column=checkDate).value,
                'CheckNumber':self.sh.cell(row=i,column=checkNumber).value,
                'CheckBank':self.sh.cell(row=i,column=checkBank).value,
                'CheckBs':self.sh.cell(row=i,column=bsCheck).value,
                'CheckUs':self.sh.cell(row=i,column=usCheck).value,
                'TransferDate':self.sh.cell(row=i,column=dateTransfer).value,
                'TransferBank':self.sh.cell(row=i,column=bankTransfer).value,
                'TransferBs':self.sh.cell(row=i,column=bsTransfer).value,
                'TransferUs':self.sh.cell(row=i,column=usTransfer).value,
                'SubtotalBs':self.sh.cell(row=i,column=subtotalBs).value,
                'SubtotalUs':self.sh.cell(row=i,column=subtotalUs).value,
                'SubtotalEqBs':self.sh.cell(row=i,column=subtotalEqBs).value,
                'TotalBs':self.sh.cell(row=i,column=totalBs).value,
            }
            if self.sh.cell(row=i,column=appNumber).value!=None and self.sh.cell(row=i,column=appNumber).value not in filtersKwords:
                reciboDeCajaTable.append(ditTable)
            i+=1
        print(pd.DataFrame(reciboDeCajaTable))
        return reciboDeCajaTable
    def CollectorToBoxTable(self):
        tableColumns=self.indexColumns['distribuidora']['Cierre de cobrador']
        tableKwords=self.kwordsRowLimits['distribuidora']['ambos']['recepcion en caja']
        upperLimit=tableKwords['superior']
        lowerLimit=tableKwords['inferior']
        
        filtersKwords=['Recepción en caja',"Efectivo","Bs."]

        cashBs=tableColumns['Bs']['efectivo']['recepcion en caja']['Bs']
        cashUs=tableColumns['Us']['efectivo']['recepcion en caja']['Us']
        cashEqBs=tableColumns['Bs']['efectivo']['recepcion en caja']['EqBs']
        checkBs=tableColumns['Bs']['Cheque']['recepcion en caja']['Bs']
        checkUs=tableColumns['Us']['Cheque']['recepcion en caja']['Us']
        checkEqBs=tableColumns['Bs']['Cheque']['recepcion en caja']['EqBs']

        dateTransfer=tableColumns['ambos']['transferencia']['recepcion en caja']['Fecha Transferencia']
        bankTransfer=tableColumns['ambos']['transferencia']['recepcion en caja']['banco']
        bsTransfer=tableColumns['Bs']['transferencia']['recepcion en caja']['Bs']
        usTransfer=tableColumns['Us']['transferencia']['recepcion en caja']['Us']
        EqBsTransfer=tableColumns['Bs']['transferencia']['recepcion en caja']['EqBs']
        totalBs=tableColumns['Bs']['total']['recepcion en caja']['Bs']
        
        upperLimit=tableKwords['superior']
        botomLimit=tableKwords['inferior']

        receiptBoxTable=[]
        i=1
        while self.sh.cell(row=i,column=11).value!=upperLimit:
            i+=1

        while self.sh.cell(row=i,column=5).value!=botomLimit:
            ditTable={
                "CashBs":self.sh.cell(row=i,column=cashBs).value,
                "CashUs":self.sh.cell(row=i,column=cashUs).value,
                "CashEqBs":self.sh.cell(row=i,column=cashEqBs).value,
                "CheckBs":self.sh.cell(row=i,column=checkBs).value,
                "CheckUs":self.sh.cell(row=i,column=checkUs).value,
                "CheckEqBs":self.sh.cell(row=i,column=checkEqBs).value,
                "TransferDate":self.sh.cell(row=i,column=dateTransfer).value,
                "TransferBank":self.sh.cell(row=i,column=bankTransfer).value,
                "TransferBs":self.sh.cell(row=i,column=bsTransfer).value,
                "TransferUs":self.sh.cell(row=i,column=usTransfer).value,
                "TransferEqBs":self.sh.cell(row=i,column=EqBsTransfer).value,
                "TotalBs":self.sh.cell(row=i,column=totalBs).value, 
            }
            if self.sh.cell(row=i,column=cashBs).value!=None and self.sh.cell(row=i,column=cashBs).value not in filtersKwords:
                receiptBoxTable.append(ditTable)
            i+=1
        #print(pd.DataFrame(receiptBoxTable))
        return receiptBoxTable
def scrap_CierreCobrador():
    print(os.path.join(get_current_path()))
    cierreCobradorFiles=os.listdir(os.path.join(get_current_path(),"Cierres de Cobrador","formatoxlsx"))
    for file in cierreCobradorFiles:
        if file.endswith(".xlsx"):
            scob=scraperCierreCobrador(file)
            q=scob.ClientToCollectorTable()
            p=scob.CollectorToBoxTable()
if __name__ == "__main__":
    scrap_CierreCobrador()
            