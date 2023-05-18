from downloadFiles import donloadSgv,downloadCollectorClosing
from utils import delete_xlsFiles, get_current_path, convert_xls,loginInfo,get_templatesSap
from scrapXlsxFiles import scrapCierresDeCaja
from scrapyCierreCobrador import scrap_CierreCobrador
from getMatrix import makeFinalTemplate
from sap import superTable,tableTransSap,insertDataToJsonAg
import os
def main():
    loginData=loginInfo()
    if loginData['flags']['flow']!="PROCESAR":
        delete_xlsFiles()
        donloadSgv(loginData)
    if loginData['flags']['flow']!="DESCARGAR":
        print("---------------------PROCESANDO ARCHIVOS...")
        boxClosingFolder=os.path.join(get_current_path(),"Cierres de Caja")
        convert_xls(boxClosingFolder)
        scrapCierresDeCaja()
        collectorClosingFolder=os.path.join(get_current_path(),"Cierres de Cobrador")
        convert_xls(collectorClosingFolder)
        scrap_CierreCobrador()
        #get_templatesSap(loginData['dates'])
        superTable(loginData)
        tableTransSap(loginData)
        insertDataToJsonAg(loginData)
        makeFinalTemplate()
        print("\n---------------------PROCESO FINALIZADO---------------------\n")
if __name__ == "__main__":
    main()