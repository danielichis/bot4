# playwright 
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    # Open new page
    page = context.new_page()
    # Go to https://www.borrd.com/
    page.goto("http://sgv.grupo-venado.com/venado/login.jsf")

    
    # Click text=Login

    page.get_by_placeholder("Usuario").click()
    page.get_by_placeholder("Usuario").fill("BOT.ADMINISTRACION.LP")
    page.get_by_placeholder("Contraseña").click()
    page.get_by_placeholder("Contraseña").fill("venadobot")
    page.get_by_role("button", name="Iniciar Sesión").click()
    page.wait_for_load_state()
    page.get_by_role("link", name="  Cobranza").click()
    page.get_by_role("link", name=" Cierres de Caja").click()
    page.wait_for_load_state()
    page.get_by_placeholder("Fecha inicial").click()
    
    page.get_by_role("row", name="30 31 1 2 3 4 5").get_by_role("cell", name="1").click()
    page.get_by_placeholder("Fecha final").click()
    page.get_by_role("row", name="30 31 1 2 3 4 5").get_by_role("cell", name="5").click()
    page.pause()
    # page.get_by_role("link", name="2").click()
    # page.get_by_role("link", name="3").click()
    # page.get_by_role("link", name="1").click()
    with page.expect_download() as download_info:
        page.locator("//td[contains(text(),'52563')]//parent::tr//a[3]/i[@class='fa fa-download']").click()
        #page.locator("//*[@id='tooltips-link']/a[3]/i").click()
        #page.get_by_role("cell", name="   Descargar EXCEL    ").get_by_role("link", name="").first.click()
        page1 = download_info.value
    download = download_info.value
    # Save downloaded file somewhere
    download.save_as("excel.xls")

