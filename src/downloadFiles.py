# playwright 
from playwright.sync_api import sync_playwright
from datetime import datetime
from datetime import timedelta
import pandas as pd
from pathlib import Path
import openpyxl
import os
import json
import time
from utils import get_current_path,pathsProyect
from doomDirections import sgvPaths
import locale
paths=pathsProyect()
def login_page(user):
    page.locator("[placeholder=\"Usuario\"]").click()
    page.locator("[placeholder=\"Usuario\"]").fill(user["user"])
    page.locator("[placeholder=\"Contraseña\"]").click()
    page.locator("[placeholder=\"Contraseña\"]").fill(user["password"])
    page.locator("input:has-text(\"Iniciar Sesión\")").click()
    page.wait_for_load_state()

def goto_bills():
    #page.pause()
    page.locator("a:has-text(\"Cobranza\")").first.click()
    page.locator("a:has-text(\"Cierres de Caja\")").first.click()
    page.wait_for_load_state()

def set_day(dExcel,cssDate):
    if cssDate=="input#startDate":
        dates=[x for x in page.query_selector_all("div:nth-child(10) div.datepicker-days tbody td[class*='day']")]
    elif cssDate=="input#endDate":
        dates=[x for x in page.query_selector_all("div:nth-child(11) div.datepicker-days tbody td[class*='day']")]
    else:
        return
    for d in dates:
        if int(d.inner_text())==int(dExcel.strftime("%d")):
            d.click()
            break

def download_file(pathFile,cssSelector,row):
    metadataFile={"path":pathFile,"descargado":"","name":os.path.splitext(os.path.basename(pathFile))[0],"retries":0}
    retries=2
    intentos=0
    donwnload=False
    while intentos<=retries and donwnload==False:
        try:
            with page.expect_download(timeout=10000) as download_info:
                row.query_selector(cssSelector).click(timeout=3000)
            download = download_info.value
            download.save_as(pathFile)
            donwnload=True
            print(f"Descargado {pathFile}")
            metadataFile["descargado"]="OK"
        except Exception as e:
            print(f"{e}\n Reitentanto...wait 2s")
            time.sleep(2)
        intentos+=1
    
    if donwnload==False:
        print(f"Error al descargar {pathFile}")
        metadataFile["descargado"]="ERROR"

    metadataFile["retries"]=intentos
    xlsFilesList.append(metadataFile)
    
def tableCashClosing(user):
    table=[]
    time.sleep(1)
    page.wait_for_load_state()
    headersTable=[x.inner_text() for x in page.query_selector_all("table#cashierClosings thead th")]
    rows=page.query_selector_all("table#cashierClosings tbody tr")
    rowsTd=page.query_selector_all("table#cashierClosings tbody tr td")
    if len(rowsTd)==1:
        print("No hay cierres de caja")
        return
    global xlsFilesList
    for row in rows:
        xlsFilesList=[]
        if len(row.query_selector_all("a"))==7:
            tipe="distribuidora"
        elif len(row.query_selector_all("a"))==5:
            tipe="agencia"
        else:
            tipe="otro"
        xpathArceoCajaBs="a[data-original-title='Arqueo de Caja Bs. EXCEL']"
        xpathArceoCajaUs="a[data-original-title='Arqueo de Caja $us. EXCEL']"
        xpathFirstExcel="a[data-original-title='Descargar EXCEL']"
        fields=[y.inner_text() for y in row.query_selector_all("td")]
        cashCode=fields[0]
        recau=user['recaudadora']
        if tipe=="distribuidora":
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{recau}_{cashCode}_arceoCajaBs_dist.xls"),xpathArceoCajaBs,row)
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{recau}_{cashCode}_arceoCajaUs_dist.xls"),xpathArceoCajaUs,row)
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{recau}_{cashCode}_firstExcel_dist.xls"),xpathFirstExcel,row)
        elif tipe=="agencia":
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{recau}_{cashCode}_arceoCajaBs_ag.xls"),xpathArceoCajaBs,row)
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{recau}_{cashCode}_arceoCajaUs_ag.xls"),xpathArceoCajaUs,row)
            
        else:
            pass 
        
        rowDict={
            headersTable[0]:fields[0],
            headersTable[1]:fields[1],
            headersTable[2]:fields[2],
            headersTable[3]:fields[3],
            headersTable[4]:fields[4],
            headersTable[5]:fields[5],
            headersTable[6]:fields[6],
            headersTable[7]:fields[7],
            headersTable[8]:fields[8],
            headersTable[9]:tipe,
            "xlsFilesList":xlsFilesList,
            "Recaudadora":user["recaudadora"]
        }
        table.append(rowDict)
    return table
def evaluate_month(monthdate_obj,dExcel,cssDate):
    tday=dExcel.strftime("%B %Y")
    if cssDate=="input#startDate":
        prevSelector="div.datepicker-days th.prev"
    elif cssDate=="input#endDate":
        prevSelector="div:nth-child(11) div.datepicker-days th.prev"
    if monthdate_obj.strftime("%B %Y")==tday:
        #print("same month")
        set_day(dExcel,cssDate)
        return True
    elif monthdate_obj<dExcel:
        #print("next month")
        page.query_selector("div.datepicker-days th.next").click()
        return False
        #monthdate=w.find_element(By.CSS_SELECTOR,"div.datepicker-days th.datepicker-switch").text
    elif monthdate_obj>dExcel:
        #print("previous month")
        page.query_selector(prevSelector).click()
        return False
def found_date(dExcel,cssDate):
    page.query_selector(cssDate).click()
    if cssDate=="input#startDate":
        monthSelector="div.datepicker-days th.datepicker-switch"
        monthdate=page.query_selector("body > div:nth-child(10) > div.datepicker-days > table > thead > tr:nth-child(1) > th.datepicker-switch").inner_text()
        monthdate=monthdate.replace("Septiembre","Setiembre")
        monthdate_obj=datetime.strptime(monthdate,"%B %Y")
    elif cssDate=="input#endDate":
        monthSelector="div:nth-child(11) div.datepicker-days th.datepicker-switch"
        page.wait_for_selector("body > div:nth-child(11) > div.datepicker-days > table > thead > tr:nth-child(1) > th.datepicker-switch")
        monthdate=page.query_selector("body > div:nth-child(11) > div.datepicker-days > table > thead > tr:nth-child(1) > th.datepicker-switch").inner_text()
        monthdate=monthdate.replace("Septiembre","Setiembre")
        #print(monthdate)
        monthdate_obj=datetime.strptime(monthdate,"%B %Y")
   
    dateNotfound=True
    while dateNotfound:
        if evaluate_month(monthdate_obj,dExcel,cssDate):
            dateNotfound=False
            #print("date found")
        else:
            monthdate=page.query_selector(monthSelector).inner_text()
            monthdate=monthdate.replace("Septiembre","Setiembre")
            monthdate_obj=datetime.strptime(monthdate,"%B %Y")

def in_folder(nameFolder):
    folderParent = paths.folderProyect
    #folderParent=Path(folderParent).parent
    folderParent=os.path.join(folderParent,nameFolder)
    return folderParent
def pageFilesDownloaded(ccjaList):
    boolPage=True
    for row in ccjaList:
        for file in row["xlsFilesList"]:
            if file['descargado']!="OK":
                boolPage=False
                break
    return boolPage
def downloadCcaj(loginInfo,user):
    print("DESCARGANDO ARCHIVOS DE CIERRES DE CAJA...")
    dinit=loginInfo['dates']['dInit']
    dEnd=loginInfo['dates']['dEnd']
    locale.setlocale(locale.LC_TIME, '')
    globalList=[]
    goto_bills()
    found_date(dinit,"input#startDate")
    time.sleep(1)
    found_date(dEnd,"input#endDate")
    page.wait_for_selector("table#cashierClosings td")
    page.evaluate('window.scrollBy(0, 200)')
    #time.sleep(2)
    n=1
    i=0
    pageBool=False
    maxRetriesPage=3
    retryPage=0
    while n>0:
        print(f"------pagina : {i+1}")
        ccjaList=tableCashClosing(user)
        pageBool=pageFilesDownloaded(ccjaList)
        if ccjaList:
            globalList.extend(ccjaList)        
        #time.sleep(2)
        if pageBool or retryPage==maxRetriesPage:
            page.query_selector("[id='cashierClosings_next'] a").click()
            i=i+1
            retryPage+=1
        n=len(page.query_selector_all("li[class='paginate_button next'] a"))

    listofFilesData={}
    listofFilesData["data"]=globalList
    with open(paths.jsonCcaj, "r") as json_file: 
        data=json.load(json_file)
    if data['data']:
        print("data already exists, acumulando...")
        data['data'].extend(globalList)
    else:
        print("data not exists, creating...")
        data['data']=globalList

    with open (paths.jsonCcaj,"w") as json_file:
        json_file.write(json.dumps(data,indent=4))

def collectorClosingFrame():
    sgvp=sgvPaths()
    #page.locator(sgvp.collections['XPATH']).first.click()
    page.wait_for_load_state('load')
    page.wait_for_load_state('networkidle')
    page.locator(sgvp.collectorClosingBtn['XPATH']).first.click()
    page.wait_for_load_state()

def cashOutFrame():
    sgvp=sgvPaths()
    try:
        page.locator(sgvp.collections['XPATH']).first.click()
        page.wait_for_load_state('load')
        page.wait_for_load_state('networkidle')
        page.locator(sgvp.cashOut.cashOutBtn['XPATH']).click()
        page.wait_for_load_state()
    except:
        page.locator(sgvp.collections['XPATH']).first.click()
        page.wait_for_load_state('load')
        page.wait_for_load_state('networkidle')
        page.locator(sgvp.cashOut.cashOutBtn['XPATH']).click()
        page.wait_for_load_state()
        
def tableCollectorClosing(user):
    sgvp=sgvPaths()
    closingTable=[]
    recaud=user['recaudadora']
    page.wait_for_selector(sgvp.collectorClosing.dailyClosingCollectorTable['CSS'])
    time.sleep(2)
    closingTableFrame=page.query_selector_all(sgvp.collectorClosing.dailyClosingCollectorTable['CSS'])
    closingTableFrameTd=page.query_selector_all(sgvp.collectorClosing.dailyClosingCollectorTableTd['CSS'])
    if len(closingTableFrameTd)==1:
        print("No hay cierre de cobrador")
        return
    for row in closingTableFrame:
        date=row.query_selector("//td[3]").inner_text().replace("/","")
        amount=str(row.query_selector("//td[6]").inner_text()).replace(",","")
        uniqueId=row.query_selector("//td[5]").inner_text()+"_"+date+"_"+amount
        nameFile=f"{recaud}_{uniqueId}.xls"
        closingTableDict={
            "codigo":row.query_selector("//td[1]").inner_text(),
            "Recibo":row.query_selector("//td[2]").inner_text(),
            "Fecha de Creacion":row.query_selector("//td[3]").inner_text(),
            "Correspondiente al":row.query_selector("//td[4]").inner_text(),
            "Cobrador":row.query_selector("//td[5]").inner_text(),
            "Total (Bs)":row.query_selector("//td[6]").inner_text(),
            "Estado":row.query_selector("//td[7]").inner_text(),
            "Nombre del archivo":nameFile,
            "UniqueID":uniqueId,
            "xlsFilesList":xlsFilesList   
        }
        #row.query_selector(sgvp.collectorClosing.excelDonwloadBtn["CSS"]).click(timeout=5000)
        
        pathFile=os.path.join(in_folder("Cierres de cobrador"),nameFile)
        download_file(pathFile,sgvp.collectorClosing.excelDonwloadBtn["CSS"],row)
        closingTable.append(closingTableDict)
    return closingTable
def downloadCollectorClosing(loginInfo,user):
    print("DESCARGANDO ARCHIVOS DE CIERRES DE COBRADOR...")
    sgvp=sgvPaths()
    configInfo=loginInfo['dates']
    dinit=configInfo["dInit"]
    dEnd=configInfo["dEnd"]
    locale.setlocale(locale.LC_TIME, '')
    globalList2=[]
    collectorClosingFrame()
    found_date(dinit,"input#startDate")
    time.sleep(1)
    found_date(dEnd,"input#endDate")
    page.wait_for_load_state("networkidle")
    page.evaluate('window.scrollBy(0, 200)')
    i=0
    n=1
    pageBool=False
    maxRetriesPage=3
    retryPage=0
    while n>0:
        print(f"-----pagina :{i+1}")
        ccobList=tableCollectorClosing(user)
        pageBool=pageFilesDownloaded(ccobList)
        if ccobList:
            globalList2.extend(ccobList)
        if pageBool or retryPage==maxRetriesPage:
            i=i+1
            page.query_selector("[id='dailyClosings_next'] a").click()
            retryPage=+1
        n=len(page.query_selector_all("li[class='paginate_button next'] a"))

    listofFilesData={}
    listofFilesData["data"]=globalList2
    with open(paths.jsonCcob, "r") as outfile: 
        data=json.load(outfile)
    if data['data']:
        data['data'].extend(globalList2)
    else:
        data['data']=globalList2
    with open (paths.jsonCcob,"w") as outfile:
        outfile.write(json.dumps(data,indent=4))

def cashOutTable(user):
    sgvp=sgvPaths()
    eTable=page.wait_for_selector(sgvp.cashOut.CashOutTable['CSS'])
    rows=eTable.query_selector_all("tbody tr")
    count_row=eTable.query_selector_all("tbody tr td")
    if len(count_row)==1:
        print("No hay Salidas de Efectivo para este rango de fechas")
        #page.pause()
        return
    cashOutList=[]
    for row in rows:
        code=row.query_selector("td:nth-child(1)").inner_text()
        date=row.query_selector("td:nth-child(2)").inner_text()
        manager=row.query_selector("td:nth-child(3)").inner_text()
        TotalBs=row.query_selector("td:nth-child(4)").inner_text()
        typeEnterprise=row.query_selector("td:nth-child(5)").inner_text()
        state=row.query_selector("td:nth-child(6)").inner_text()

        cashOutDict={
            "Código":code,
            "Fecha":date,
            "Encargado":manager,
            "Total Bs.":TotalBs,
            "Tipo":typeEnterprise,
            "Estado":state,
            "Recaudadora":user["recaudadora"]
        }
        cashOutList.append(cashOutDict)
    return cashOutList
def get_outOffCashSgv(loginInfo,user):
    print("downloading SALIDAS DE EFECTIVO")
    sgvp=sgvPaths()
    configInfo=loginInfo['dates']
    dinit=configInfo["dInit"]
    dinit=dinit-timedelta(days=5)
    dEnd=configInfo["dEnd"]
    dEnd=dEnd+timedelta(days=5)
    locale.setlocale(locale.LC_TIME, '')
    globalList3=[]
    cashOutFrame()
    found_date(dinit,"input#startDate")
    time.sleep(1)
    found_date(dEnd,"input#endDate")
    page.wait_for_load_state("networkidle")
    page.evaluate('window.scrollBy(0, 200)')
    #page.pause()
    i=0
    n=1
    while n>0:
        time.sleep(1)
        print(f"-----pagina :{i+1}")
        cashTable=cashOutTable(user)
        if cashTable:
            globalList3.extend(cashTable)
        i=i+1
        page.query_selector("[id='cashOuts_next'] a").click()
        n=len(page.query_selector_all("li[class='paginate_button next'] a"))
    return globalList3
def saveCashoutSgv(cashoutList):
    with open(paths.jsonCashOut, "w") as outfile: 
        outfile.write(json.dumps(cashoutList,indent=4))
    df=pd.DataFrame(cashoutList)
    df.to_csv(paths.csvCashOut,index=False,encoding="utf-8-sig",sep=";")

def donloadSgv(loginInfo):
    print("--------------------DESCARGANDO ARCHIVOS DEL SGV----------------------")
    #context=browser.new_context(record_video_dir="videos/")
    global page
    users=loginInfo["users"]
    allCashOuts=[]
    for key,user in users.items():
        p=sync_playwright().start()
        browser=p.chromium.launch(headless=False)
        page=browser.new_page(accept_downloads=True)    
        print(f"-----------RECAUDADORA: {key}-----------------")
        page.goto("http://sgv.grupo-venado.com/venado/login.jsf")
        login_page(user)  
        downloadCcaj(loginInfo,user)
        downloadCollectorClosing(loginInfo,user)
        cashOutAccount=get_outOffCashSgv(loginInfo,user)
        if cashOutAccount:
            allCashOuts.extend(cashOutAccount)
        page.close()
        browser.close()
        p.stop()
        
    saveCashoutSgv(allCashOuts)
    
if __name__ == "__main__":
    donloadSgv()