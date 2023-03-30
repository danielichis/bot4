from downloadFiles import donloadSgv,downloadCollectorClosing
from utils import delete_xlsFiles, get_current_path, convert_xls,loginInfo
from scrapXlsxFiles import scrapCierresDeCaja
from scrapyCierreCobrador import scrap_CierreCobrador
import os
def main():
    loginData=loginInfo()
    if loginData['flags']['cumulative']=="NO":
        delete_xlsFiles()
    if loginData['flags']['flow']!="PROCESAR":
        donloadSgv(loginData)
    if loginData['flags']['flow']!="DESCARGAR":
        boxClosingFolder=os.path.join(get_current_path(),"Cierres de Caja")
        convert_xls(boxClosingFolder)
        scrapCierresDeCaja()
        collectorClosingFolder=os.path.join(get_current_path(),"Cierres de Cobrador")
        convert_xls(collectorClosingFolder)
        scrap_CierreCobrador()
if __name__ == "__main__":
    main()