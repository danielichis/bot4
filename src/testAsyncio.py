from playwright.sync_api import sync_playwright

p=sync_playwright().start()
browser=p.chromium.launch(headless=False)
context=browser.new_context(record_video_dir="videos/")
page=browser.new_page(accept_downloads=True)
page.goto("http://sgv.grupo-venado.com/venado/login.jsf")

#close browser
page.close()
browser.close()
p.stop()
