# playwright 
from playwright.sync_api import sync_playwright
from datetime import datetime
import pandas as pd
from pathlib import Path
import openpyxl
import os
import json
import time
from utils import get_current_path
from doomDirections import sgvPaths
import locale
def init_page():
    # Go to https://www.borrd.com/
    page.goto("http://sgv.grupo-venado.com/venado/login.jsf")
    #page.pause()
    page.locator("[placeholder=\"Usuario\"]").click()
    page.locator("[placeholder=\"Usuario\"]").fill("BOT.ADMINISTRACION.LP")
    page.locator("[placeholder=\"Contraseña\"]").click()
    page.locator("[placeholder=\"Contraseña\"]").fill("venadobot")
    page.locator("input:has-text(\"Iniciar Sesión\")").click()
    page.wait_for_load_state()

def goto_bills():
    #page.pause()
    page.locator("a:has-text(\"Cobranza\")").first.click()
    page.locator("a:has-text(\"Cierres de Caja\")").first.click()
    page.wait_for_load_state()

def set_day(dExcel,cssDate):
    if cssDate=="input#startDate":
        dates=[x for x in page.query_selector_all("div:nth-child(10) div.datepicker-days tbody td[class='day']")]
    elif cssDate=="input#endDate":
        dates=[x for x in page.query_selector_all("div:nth-child(11) div.datepicker-days tbody td[class='day']")]
    else:
        return
    for d in dates:
        if int(d.inner_text())==int(dExcel.strftime("%d")):
            d.click()
            break

def download_file(nameFile,cssSelector,row):
    with page.expect_download() as download_info:
        row.query_selector(cssSelector).click()
    download = download_info.value
    download.save_as(nameFile)

def tableCashClosing():
    table=[]
    time.sleep(1)
    page.wait_for_load_state()
    headersTable=[x.inner_text() for x in page.query_selector_all("table#cashierClosings thead th")]
    rows=page.query_selector_all("table#cashierClosings tbody tr")
    print(len(rows))
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
        if tipe=="distribuidora":
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{cashCode}_arceoCajaBs_dist.xls"),xpathArceoCajaBs,row)
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{cashCode}_arceoCajaUs_dist.xls"),xpathArceoCajaUs,row)
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{cashCode}_firstExcel_dist.xls"),xpathFirstExcel,row)
        elif tipe=="agencia":
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{cashCode}_arceoCajaBs_ag.xls"),xpathArceoCajaBs,row)
            download_file(os.path.join(in_folder("Cierres de Caja"),f"{cashCode}_arceoCajaUs_ag.xls"),xpathArceoCajaUs,row)
            
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
            "xlsFilesList":xlsFilesList
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
        print("same month")
        set_day(dExcel,cssDate)
        return True
    elif monthdate_obj<dExcel:
        print("next month")
        page.query_selector("div.datepicker-days th.next").click()
        return False
        #monthdate=w.find_element(By.CSS_SELECTOR,"div.datepicker-days th.datepicker-switch").text
    elif monthdate_obj>dExcel:
        print("previous month")
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
        print(monthdate)
        monthdate_obj=datetime.strptime(monthdate,"%B %Y")
   
    dateNotfound=True
    while dateNotfound:
        if evaluate_month(monthdate_obj,dExcel,cssDate):
            dateNotfound=False
            print("date found")
        else:
            monthdate=page.query_selector(monthSelector).inner_text()
            monthdate=monthdate.replace("Septiembre","Setiembre")
            monthdate_obj=datetime.strptime(monthdate,"%B %Y")

def in_folder(nameFolder):
    folderParent = os.getcwd()
    #folderParent=Path(folderParent).parent
    folderParent=os.path.join(folderParent,nameFolder)
    return folderParent

    

def downloadSgv():
    wb=openpyxl.load_workbook(os.path.join(get_current_path(),"config.xlsx"))
    ws=wb["login"]
    dinit=ws["B2"].value
    dEnd=ws["B3"].value
    locale.setlocale(locale.LC_TIME, '')
    globalList=[]
    with sync_playwright() as p:
        global browser,context,page
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
                # Open new page
        page = context.new_page()
        init_page()
        goto_bills()
        found_date(dinit,"input#startDate")
        time.sleep(1)
        found_date(dEnd,"input#endDate")
        page.wait_for_url("http://sgv.grupo-venado.com/venado/cashierclosings/cashierclosing_list.jsf")
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        page.mouse.wheel(0,200)
        #page.pause()
        #page.wait_for_selector("li[class='paginate_button next'] a")
        npaginations=len(page.query_selector_all("ul.pagination li"))-2
        #view if element is clickable or not
        for i in range(npaginations):
            print(f"page{i+1}")
            globalList.extend(tableCashClosing())
            #  time.sleep(3)
            page.query_selector("[id='cashierClosings_next'] a").click()
        listofFilesData={}
        listofFilesData["data"]=globalList
        with open(r'src\target\CashClosingInfo.json', "w") as outfile: 
            json.dump(listofFilesData, outfile,indent=4)
        #page.pause()

def collectorClosingFrame():
    sgvp=sgvPaths()
    page.locator(sgvp.collections['XPATH']).first.click()
    page.locator(sgvp.collectorClosingBtn['XPATH']).first.click()
    page.wait_for_load_state()

def tableCollectorClosing():
    sgvp=sgvPaths()
    closingTable=[]
    page.wait_for_selector(sgvp.collectorClosing.dailyClosingCollectorTable['CSS'])
    time.sleep(2)
    closingTableFrame=page.query_selector_all(sgvp.collectorClosing.dailyClosingCollectorTable['CSS'])
    for row in closingTableFrame:
        date=row.query_selector("//td[3]").inner_text().replace("/","")
        amount=str(row.query_selector("//td[6]").inner_text()).replace(",","")
        uniqueId=row.query_selector("//td[5]").inner_text()+"_"+date+"_"+amount
        nameFile=f"{uniqueId}.xls"
        closingTableDict={
            "codigo":row.query_selector("//td[1]").inner_text(),
            "Recibo":row.query_selector("//td[2]").inner_text(),
            "Fecha de Creacion":row.query_selector("//td[3]").inner_text(),
            "Correspondiente al":row.query_selector("//td[4]").inner_text(),
            "Cobrador":row.query_selector("//td[5]").inner_text(),
            "Total (Bs)":row.query_selector("//td[6]").inner_text(),
            "Estado":row.query_selector("//td[7]").inner_text(),
            "Nombre del archivo":nameFile,
            "UniqueID":uniqueId
        }
        #row.query_selector(sgvp.collectorClosing.excelDonwloadBtn["CSS"]).click(timeout=5000)
        
        pathFile=os.path.join(in_folder("Cierres de cobrador"),nameFile)
        download_file(pathFile,sgvp.collectorClosing.excelDonwloadBtn["CSS"],row)
        closingTable.append(closingTableDict)
    return closingTable
def downloadCollectorClosing():
    pass
    sgvp=sgvPaths()
    wb=openpyxl.load_workbook(os.path.join(get_current_path(),"config.xlsx"))
    ws=wb["login"]
    dinit=ws["B2"].value
    dEnd=ws["B3"].value
    locale.setlocale(locale.LC_TIME, '')
    globalList2=[]
    with sync_playwright() as p:
        global browser,context,page
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
                # Open new page
        page = context.new_page()
        init_page()
        collectorClosingFrame()
        found_date(dinit,"input#startDate")
        time.sleep(1)
        found_date(dEnd,"input#endDate")
        page.wait_for_load_state("networkidle")
        page.mouse.wheel(0,200)
        #page.pause()
        #page.wait_for_selector("li[class='paginate_button next'] a")
        npaginations=len(page.query_selector_all("ul.pagination li"))-2
        #view if element is clickable or not
        i=0
        n=len(page.query_selector_all("li[class='paginate_button next'] a"))
        print(n)
        while n>0:
            n=len(page.query_selector_all("li[class='paginate_button next'] a"))
            print(f"page{i+1}")
            globalList2.extend(tableCollectorClosing())
            i=i+1
            page.query_selector("[id='dailyClosings_next'] a").click()

        listofFilesData={}
        listofFilesData["data"]=globalList2
        with open(r'src\target\CollectorClosingFilesDonwload.json', "w") as outfile: 
            json.dump(listofFilesData, outfile,indent=4)
        #page.pause()
if __name__ == "__main__":
    downloadSgv()
    downloadCollectorClosing()