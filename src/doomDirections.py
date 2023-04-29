class sgvPaths:
    def __init__(self):
        self.Calendar = Calendar()
        self.boxClosing = BoxClosing()
        self.collectorClosing = CollectorClosing()
        self.cashOut = CashOut()
        self.collections = {"XPATH": "//span[contains(text(),'Cobranza')]", "CSS": ""}
        self.collectorClosingBtn={"XPATH": "//span[contains(text(),'Cierres de Cobradores')]", "CSS": ""}

class BoxClosing:
    def __init__(self):
        self.dailyClosingBoxTable = {"XPATH": "//table[@id='dailyClosings']/tbody", "CSS": "table#dailyClosings tbody"}
        self.rowTable = {"XPATH": "//table[@id='dailyClosings']/tbody/tr", "CSS": "table#dailyClosings tbody tr"}
        self.excelDonwloadBtn = {"XPATH": "//i[@class='fa fa-download']", "CSS": ""}
class Calendar:
    def __init__(self):
        self.initDatePicker = {"XPATH": "//div[@id='startDate-datepicker']//i[@class='fa fa-calendar']", "CSS": ""}
        self.endDatePicker = {"XPATH": "//div[@id='endDate-datepicker']//i[@class='fa fa-calendar']", "CSS": ""}

class CollectorClosing:
    def __init__(self):
        self.dailyClosingCollectorTable = {"XPATH": "//table[@id='dailyClosings']/tbody/tr", "CSS": "table#dailyClosings tbody tr"}
        self.dailyClosingCollectorTableTd={"XPATH": "//table[@id='dailyClosings']/tbody/tr/td", "CSS": "table#dailyClosings tbody tr td"}
        self.rowTable = {"XPATH": "//table[@id='dailyClosings']/tbody/tr", "CSS": "table#dailyClosings tbody tr"}
        self.excelDonwloadBtn = {"XPATH": "", "CSS": "a[data-original-title='Descargar EXCEL']"}

class CashOut:
    def __init__(self):
        self.cashOutBtn = {"XPATH": "//span[contains(text(),'Salida de Efectivo')]", "CSS": ""}
        self.CashOutTable = {"XPATH": "//table[@id='cashOuts']", "CSS": "table#cashOuts"}
