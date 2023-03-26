from downloadFiles import downloadSgv,downloadCollectorClosing
from utils import delete_xlsFiles, get_current_path, convert_xls
from scrapXlsxFiles import scrapCierresDeCaja
from scrapyCierreCobrador import scrap_CierreCobrador
import os
def main():
    delete_xlsFiles(get_current_path())
    downloadSgv()
    boxClosingFolder=os.path.join(get_current_path(),"Cierres de Caja")
    convert_xls(boxClosingFolder)
    downloadCollectorClosing()
    scrapCierresDeCaja()
    collectorClosingFolder=os.path.join(get_current_path(),"Cierres de Cobrador")
    convert_xls(collectorClosingFolder)
    scrap_CierreCobrador()
if __name__ == "__main__":
    main()