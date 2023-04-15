import subprocess
import os
from pathlib import Path
from openpyxl import load_workbook
import win32com.client
import time

changeTheDateInicio = None
changeTheDateFin = None
listOfAccounts = []


def add0(num):
  num_str = str(num)
  if len(num_str) == 1:
    return '0' + num_str
  return num_str

def superTable():
    currentPath = os.getcwd()
    SAPinfoPath = os.path.join(currentPath, "SAPinfo")
    # currentPathParentFolder = Path(currenPath).parent

    configXlsx=os.path.join(currentPath,"config.xlsx")
    wb = load_workbook(configXlsx)
    wsConfig = wb['login']            

    login = {'SAPPath': wsConfig['B16'].value,
            'user': wsConfig['B17'].value,
                'psw': wsConfig['B18'].value,
                'environment': wsConfig['B19'].value,
                'layout': wsConfig['B20'].value,
                'fechaInicio': wsConfig['B2'].value,
                'fechaFin': wsConfig['B3'].value
                }

    wsAccounts = wb['Hoja1']
    for i in range(2, wsAccounts.max_row+1):
        accountCell = wsAccounts[f'A{i}'].value
        accountCell = str(accountCell)
        accountCell = accountCell.replace(" ", "")
        if accountCell != None and accountCell != "":
            listOfAccounts.append(accountCell)

    wb.close()

    layout = login['layout']
    layout = layout.replace(" ","")

    if login['fechaInicio'] != None:
        changeTheDateInicio = True
    if login['fechaFin'] != None:
        changeTheDateFin = True

    proc = subprocess.Popen([login['SAPPath'], '-new-tab'])
    time.sleep(2)
    try: 
        sapGuiAuto = win32com.client.GetObject('SAPGUI')
    except:
        proc.kill()
        time.sleep(2)
        proc = subprocess.Popen([login['SAPPath'], '-new-tab'])
        time.sleep(2)
        sapGuiAuto = win32com.client.GetObject('SAPGUI')

    application = sapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection(login['environment'], True)
    session = connection.Children(0)

    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = login['user']
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = login['psw']
    session.findById("wnd[0]").sendVKey(0)

    session.EndTransaction()
    session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
    session.findById("wnd[0]").sendVKey(0)

    fi = login['fechaInicio']
    ff = login['fechaInicio']

    fi = f"{add0(fi.day)}.{add0(fi.month)}.{add0(fi.year)}"
    ff = f"{add0(ff.day)}.{add0(ff.month)}.{add0(ff.year)}"

    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = fi
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = ff
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 10
    session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()
    for i,j in enumerate(listOfAccounts):
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = j
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = i + 1

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = SAPinfoPath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{fi} a {ff}.txt"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 86
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


if __name__ == '__main__':
    superTable()
    print('GAAAAAAAAAAAAAAAAAAAAAAAAAA')