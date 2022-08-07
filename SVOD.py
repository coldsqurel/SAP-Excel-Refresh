import xlwings as xw
app = xw.App(visible=True, add_book=False)
wb = app.books.open(r"G:\Новое ОКК и УР\Мониторинг\Август 2022\СВОД Август.xlsm")
vba_macro = wb.macro("Macros")
vba_macro()
wb.save()
wb.close()
app.quit()
