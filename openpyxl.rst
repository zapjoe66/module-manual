Openpyxl
========
Opening Excel Documents with OpenPyXL
-------------------------------------
>>> import openpyxl
>>> wb = openpyxl.load_workbook('example.xlsx')  #get a workbook

Getting Sheets from the Workbook
--------------------------------
>>> wb.get_sheet_names()
['Sheet1', 'Sheet2', 'Sheet3']
>>> sheet = wb.get_sheet_by_name('Sheet3')  # get a Worksheet object
>>> sheet.title   #get the sheet name
'Sheet3'
>>> anotherSheet = wb.active   #The active sheet is the sheet that’s on top when the workbook is opened in Excel.
>>> sheet.max_row   #get the max row of the sheet
7
>>> sheet.max_column  #get the max column of the sheet
3

Getting Cells from the Sheets
-----------------------------
>>> sheet = wb.get_sheet_by_name('Sheet1')
>>> sheet['A1'] #get the cell object
<Cell Sheet1.A1>
>>> sheet['A1'].value   #get the cell value
>>> sheet['A1'].row     #get the row of the cell
>>> sheet['A1'].column   #get the column of the cell
>>> sheet['A1'].coordinate   #get the coordinate of the cell
'A1'
>>> sheet.cell(row=1, column=2)  #Specifying a column by letter can be tricky to program, especially because after column Z, the columns start by using two letters: AA, AB, AC, and so on. As an alternative, you can also get a cell using the sheet’s cell() method and passing integers for its row and column keyword arguments. The first row or column integer is 1, not 0. 
<Cell Sheet1.B1>

 
