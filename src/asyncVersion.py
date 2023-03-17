import asyncio
from playwright.async_api import async_playwright
import time
import os
from doomDirections import sgvPaths



async def download_file(nameFile,cssSelector,row):
    async with page.expect_download() as download_info:
        await row.query_selector(cssSelector).click()
    download = download_info.value
    download.save_as(nameFile)
    # Guardar archivo
async def download_files():
    urls = ['https://example.com/download1', 'https://example.com/download2', 'https://example.com/download3']
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        global page
        page = await browser.new_page()
        await page.goto('http://sgv.grupo-venado.com/venado/login.jsf') # Suponga que la página tiene botones para descargar los archivos
        await page.locator("[placeholder=\"Usuario\"]").click()
        await page.locator("[placeholder=\"Usuario\"]").fill("BOT.ADMINISTRACION.LP")
        await page.locator("[placeholder=\"Contraseña\"]").click()
        await page.locator("[placeholder=\"Contraseña\"]").fill("venadobot")
        await page.locator("input:has-text(\"Iniciar Sesión\")").click()
        await page.wait_for_load_state()
        tasks = []
        sgvp=sgvPaths()
        closingTable=[]
        await page.wait_for_selector(sgvp.collectorClosing.dailyClosingCollectorTable['CSS'])
        time.sleep(2)
        closingTableFrame=await page.query_selector_all(sgvp.collectorClosing.dailyClosingCollectorTable['CSS'])
        for row in closingTableFrame:
            closingTableDict={
                "codigo":row.query_selector("//td[1]").inner_text(),
                "Recibo":row.query_selector("//td[2]").inner_text(),
                "Fecha de Creacion":row.query_selector("//td[3]").inner_text(),
                "Correspondiente al":row.query_selector("//td[4]").inner_text(),
                "Cobrador":row.query_selector("//td[5]").inner_text(),
                "Total (Bs)":row.query_selector("//td[6]").inner_text(),
                "Estado":row.query_selector("//td[7]").inner_text(),
            }
            #row.query_selector(sgvp.collectorClosing.excelDonwloadBtn["CSS"]).click(timeout=5000)
            cashCode=await row.query_selector("//td[1]").inner_text()
            checker=await row.query_selector("//td[5]").inner_text()
            nameFile=f"{cashCode}_{checker}.xls"
            pathFile=os.path.join(in_folder("Cierres de cobrador"),nameFile)
            tasks.append(asyncio.create_task(download_file(page, row.query_selector(sgvp.collectorClosing.excelDonwloadBtn["CSS"]), pathFile)))
            download_file(pathFile,sgvp.collectorClosing.excelDonwloadBtn["CSS"],row)
            closingTable.append(closingTableDict)
        await asyncio.gather(*tasks)

if __name__ == '__main__':
    asyncio.run(download_files())




