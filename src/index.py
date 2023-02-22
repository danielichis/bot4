from downloadFiles import downloadSgv
from utils import delete_xlsFiles, get_current_path, convert_xls
def main():
    #delete_xlsFiles(get_current_path())
    #downloadSgv()
    convert_xls(get_current_path())   
    pass
if __name__ == "__main__":
    main()