import xlwings as xw
app = xw.App(visible=True, add_book=False)
wb = app.books.open(r"C:\Users\alexandrovn\Desktop\otchet\МО2022 таб.xlsm")
vba_macro = xw.macro("Workbook_Open")
vba_macro()
wb.save()
wb.close()
app.quit()
