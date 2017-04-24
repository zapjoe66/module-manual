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
>>> sheet['A1'] #get the cell object in A1
<Cell Sheet1.A1>
>>> sheet['A1'].value   #get the cell value
>>> sheet['A1'].row     #get the row of the cell
>>> sheet['A1'].column   #get the column of the cell
>>> sheet['A1'].coordinate   #get the coordinate of the cell
'A1'
>>> sheet.cell(row=1, column=2)  

Specifying a column by letter can be tricky to program, especially because after column Z, the columns start by using two letters: AA, AB, AC, and so on. As an alternative, you can also get a cell using the sheet’s cell() method and passing integers for its row and column keyword arguments. The first row or column integer is 1, not 0. 
<Cell Sheet1.B1>

Converting Between Column Letters and Numbers
---------------------------------------------
>>> from openpyxl.utils import get_column_letter,column_index_from_string 
>>> get_column_letter(1)
'A'
>>> column_index_from_string('AA')
27

Getting Rows and Columns from the Sheets
----------------------------------------
You can slice Worksheet objects to get all the Cell objects in a row, column, or rectangular area of the spreadsheet. Then you can loop over all the cells in the slice.

>>> sheet['A1':'C3']
((<Cell '工作表1'.A1>, <Cell '工作表1'.B1>, <Cell '工作表1'.C1>), (<Cell '工作表1'.A2>, <Cell '工作表1'.B2>, <Cell '工作表1'.C2>), (<Cell '工作表1'.A3>, <Cell '工作表1'.B3>, <Cell '工作表1'.C3>))
>>> for rowOfCellObjects in sheet['A1':'C3']:  # loop over every rows
...     for cellObj in rowOfCellObjects:  #loop every cells in one row
...         print(cellObj.coordinate, cellObj.value)
>>> tuple(sheet.columns) 
((<Cell '工作表1'.A1>, <Cell '工作表1'.A2>, <Cell '工作表1'.A3>), (<Cell '工作表1'.B1>, <Cell '工作表1'.B2>, <Cell '工作表1'.B3>), (<Cell '工作表1'.C1>, <Cell '工作表1'.C2>, <Cell '工作表1'.C3>))
>>> tuple(sheet.rows)
((<Cell '工作表1'.A1>, <Cell '工作表1'.B1>, <Cell '工作表1'.C1>), (<Cell '工作表1'.A2>, <Cell '工作表1'.B2>, <Cell '工作表1'.C2>), (<Cell '工作表1'.A3>, <Cell '工作表1'.B3>, <Cell '工作表1'.C3>))

Creating and Saving Excel Documents
-----------------------------------

>>> sheet.title = 'Spam Bacon Eggs Sheet'  #change the sheet's title to 'Spam Bacon Eggs Sheet'
>>> wb.save('example_copy.xlsx') # save a workbook

Creating and Removing Sheets
----------------------------

>>> wb = openpyxl.Workbook()   #create a new empty workbook
>>> wb.get_sheet_names()
['Sheet']
>>> wb.create_sheet()  
<Worksheet "Sheet1">

The create_sheet() method returns a new Worksheet object named SheetX, which by default is set to be the last sheet in the workbook. Optionally, the index and name of the new sheet can be specified with the index and title keyword arguments.
 
>>> wb.get_sheet_names()
['Sheet', 'Sheet1']
>>> wb.create_sheet(index=0, title='First Sheet')
<Worksheet "First Sheet">
>>> wb.get_sheet_names()
['First Sheet', 'Sheet', 'Sheet1']
>>> wb.create_sheet(index=2, title='Middle Sheet')
<Worksheet "Middle Sheet">
>>> wb.get_sheet_names()
['First Sheet', 'Sheet', 'Middle Sheet', 'Sheet1']
>>> wb.remove_sheet(wb.get_sheet_by_name('Middle Sheet'))
>>> wb.remove_sheet(wb.get_sheet_by_name('Sheet1'))
>>> wb.get_sheet_names()
['First Sheet', 'Sheet']        

Writing Values to Cells
-----------------------
>>> sheet['A1'] = 'Hello world!'  #write value into A1
>>> sheet['A1'].value
'Hello world!'

Setting the Font Style of Cells
-------------------------------
>>> from openpyxl.styles import Font
>>> wb = openpyxl.Workbook()
>>> sheet = wb.get_sheet_by_name('Sheet')
>>> italic24Font = Font(size=24, italic=True)   
>>> sheet['A1'].font = italic24Font #A cell’s style can be set by assigning the Font object to the font attribute.
>>> fontObj1 = Font(name='Times New Roman', bold=True)
>>> sheet['A1'].font = fontObj1
>>> sheet['A1'] = 'Bold Times New Roman'

Font Objects
------------
Keyword Arguments for Font

+------------------+-----------+-------------------------------------------------------+
| Keyword argument | data type | Description                                           |
+------------------+-----------+-------------------------------------------------------+
| name             | string    | The font name, such as 'Calibri' or 'Times New Roman' |
+------------------+-----------+-------------------------------------------------------+
| size             | integer   | The point size                                        |
+------------------+-----------+-------------------------------------------------------+
| bold             | boolean   | True, for bold font                                   |
+------------------+-----------+-------------------------------------------------------+
| italic           | boolean   | True, for italic font                                 |
+------------------+-----------+-------------------------------------------------------+

Formulas
--------

>>> sheet['B9'] = '=SUM(B1:B8)'

This will store =SUM(B1:B8) as the value in cell B9. This sets the B9 cell to a formula that calculates the sum of values in cells B1 to B8. 

Setting Row Height and Column Width
-----------------------------------











        
        
        
        
        
        
        
