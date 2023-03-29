from downloadFiles import downloadCcaj,downloadCollectorClosing
from utils import delete_xlsFiles, get_current_path, convert_xls,loginInfo
from scrapXlsxFiles import scrapCierresDeCaja
from scrapyCierreCobrador import scrap_CierreCobrador
import os
def main():
    loginData=loginInfo()
    if loginData['flags']['cumulative']=="NO":
        print("Borrando archivos anteriores")
        delete_xlsFiles(get_current_path())
    if loginData['flags']['flow']!="PROCESAR":
        print("Descargando archivos")
        downloadCcaj(loginData)
    if loginData['flags']['flow']!="DESCARGAR":
        print("Procesando archivos")
        boxClosingFolder=os.path.join(get_current_path(),"Cierres de Caja")
        convert_xls(boxClosingFolder)
        downloadCollectorClosing(loginData)
        scrapCierresDeCaja()
        collectorClosingFolder=os.path.join(get_current_path(),"Cierres de Cobrador")
        convert_xls(collectorClosingFolder)
        scrap_CierreCobrador()
if __name__ == "__main__":
    main()