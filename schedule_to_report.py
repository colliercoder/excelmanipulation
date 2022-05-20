class horario():
    """a class to plug in the month"""
    def __init__(self,month):
        self.month = month
    def openfiles(self):
        import xlwings as xw
        import openpyxl
        # Establish a connection to a workbook
        hours_report = xw.Book(r"C:\Users\15402\Desktop\python_hours_script\hours_report.xlsx")
        miner_schedule = xw.Book(r"C:\Users\15402\Desktop\python_hours_script\miner_schedule.xlsx")
        wb = openpyxl.load_workbook(r"C:\Users\15402\Desktop\python_hours_script\miner_schedule.xlsx")

        # Instantiate a sheet object
        schedule = miner_schedule.sheets['MUZO EMPLOYEES'+' '+str(self.month).upper()]
        extra_horas = hours_report.sheets['HORA EXTRA']
        domingo_festivo = hours_report.sheets['DOMINGO Y FESTIVO']
        recargo_nocturno = hours_report.sheets['RECARGO NOCTURNO']

        sheet = wb['MUZO EMPLOYEES' + ' ' + str(self.month).upper()]
        lista = hours_report.sheets['LISTA']