from downloadFiles import downloadSgv
from utils import delete_xlsFiles, get_current_path, convert_xls
from scrapXlsxFiles import scrapFiles
def main():
    delete_xlsFiles(get_current_path())
    downloadSgv()
    convert_xls(get_current_path())
    scrapFiles()
if __name__ == "__main__":
    main()