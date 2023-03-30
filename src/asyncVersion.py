import asyncio
from playwright.async_api import Playwright, async_playwright
from utils import loginInfo

async def download_files(playwright: Playwright) -> None:
    
    browser = await playwright.chromium.launch(headless=False)
    context = await browser.new_context(record_video_dir="videos/")
    #await context.tracing.start(screenshots=True, snapshots=True)
    #page=await context.new_page()
    page = await browser.new_page(accept_downloads=True)
    await page.goto("http://sgv.grupo-venado.com/venado/login.jsf")
    await page.locator("[placeholder=\"Usuario\"]").click()
    await page.locator("[placeholder=\"Usuario\"]").fill("BOT.ADMINISTRACION.LP")
    await page.locator("[placeholder=\"Contraseña\"]").click()
    await page.locator("[placeholder=\"Contraseña\"]").fill("venadobot")
    await page.locator("input:has-text(\"Iniciar Sesión\")").click()
    await page.wait_for_load_state()
    await page.locator("a:has-text(\"Cobranza\")").first.click()
    await page.locator("a:has-text(\"Cierres de Caja\")").first.click()
    await page.wait_for_load_state()
    print("Logged in")
    await page.wait_for_load_state('domcontentloaded')
    await page.wait_for_load_state('networkidle')
    #links = await page.query_selector_all("a.download")
    rows=await page.query_selector_all("table#cashierClosings tbody tr")
    #table#cashierClosings tbody tr
    #cajas=await page.query_selector_all("table#cashierClosings tbody tr td:nth-child(2)").all_inner_texts()
    print(f"Found {len(rows)} links")
    tasks = []
    for i, row in zip(range(len(rows)), rows):
        cierreCaja = await row.query_selector("td:nth-child(2)")
        downloadButton = await row.query_selector("a[data-original-title='Arqueo de Caja Bs. EXCEL']")
        textoCierraCaja = await cierreCaja.text_content()
        textoCierraCaja = textoCierraCaja.replace("/", "")
        #cierreCajatext=cierreCaja.inner_text()
        filename = f"archivo_{textoCierraCaja}.xls"
        #await asyncio.sleep(2)
        task = asyncio.create_task(download_file(page, downloadButton, filename))
        tasks.append(task)
    await asyncio.gather(*tasks)
    #await context.tracing.stop(path = "trace.zip")
    await browser.close()

async def download_file(page, downloadButton, filename):
    # this functions is repeting the download_info.value, suggest a fix
    #await downloadButton.click()
    #download = await page.wait_for_download()

    async with page.expect_download() as download_info:
        print(f"Downloading file {filename}")
        await downloadButton.click()
        await asyncio.sleep(5)
        link=await downloadButton.get_attribute("href")
        #print(f"element web clicking {await downloadButton.inner_html()}")
        print(f"element web clicking {link}")
        download=await download_info.value
        await asyncio.sleep(5)
        
    print(f"archivo descargado {filename}")
    print(await download.path())
    
    await download.save_as(filename)
    print(f"archivo en directorio {filename}")
    

async def main():
    async with async_playwright() as playwright:
        print("Starting downloads")
        await download_files(playwright)

asyncio.run(main())

