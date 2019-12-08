'Steps
'---------------------------------------------------------------------------
'Part 1:
'## Instructions

'* Create a script that will loop through all the stocks for one year for each run and take the following information.

 ' * The ticker symbol.

  '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The total stock volume of the stock.

'* You should also have conditional formatting that will highlight positive change in green and negative change in red.

'* The result should look as follows.

'---------------------------------------------------------------------------------------

Sub Aplha_Test()

'----------------------------------------------
'LOOK THOUROUGH ALL SHEETS
'---------------------------------------------

For Each ws in Worksheets

'--------------------------------------------------
' SET VARIABLES
'----------------------------------------------------

Dim Ticker As String
Dim Year_Change As Double
Dim Percentage_Change As Integer
Dim Total_Stock_Volume As Integer
Dim Summary_Table_Row As Integer
Dim Close_Date As Integer
Dim Open_Date As Integer

Summary_Table_Row = 2
Year_Change = 0
Percentage_change = 0

'-------------------------------------------------------------
'GRABBED THE WORKSHEETSNAMES : NOT SURE WHAT THIS IS FOR??
'----------------------------------------------------------------

'worksheetsName = ws.Name

'---------------------------------------------------------
'SET LASTROW TO LOOP THROUGH ALL Cells
'------------------------------------------------------

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'-------------------------------------------------------------------------------
'ADD THE WORD TICKER, YEAR_CHANGE, PERCENTAGE_CHANGE, AND TOTAL_STOCK TO HEADERS
'-----------------------------------------------------------------------------

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Year_Change"
ws.Cells(1, 11).Value = "Percent_Change"
ws.Cells(1, 12).Value = "Total_Stock_Volume"

'------------------------------------------------------
'LOOP THROUGH ALL <TICKER> VALUES
'-------------------------------------------------------

For i = 2 To lastrow

For j = 2 To lastcolumn

'-----------------------------------------------------------------
'wILL CHECK IF WE'RE UNDER THE SAME TICKER NAME IF NOT THEN LOOP WILL ADD NEXT VALUES
'------------------------------------------------------------------------------


'If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'-------------------------------------------------------------
'SET THE CELL VALUES
'--------------------------------------------

Ticker = Cells(i, 1).Value
Open_Date = Cells(i, 3).Value
Close_Date = Cells(i, 6). Value

'Calculate year change
Year_Change = Year_Change + Open_Date - Close_Date

'Calculate Percentage change
Percentage_Change = Percentage_Change + Open_Date / Close_Date

'Add to the Year_Change
'Year_Change = Year_Change + Cells(i, 3).Value - Cells(i, 6).Value

'Percentage change between open and closing 
'Percentage_Change = Cells(i, 3).Value / Cells(i, 6).Value

'------------------------------------------------
'PLACE PULLED VALUES TO SPECIFIC Cells
'-------------------------------------------------------

'Print the Ticker Name
ws.Range("I" & Summary_Table_Row).Value = Ticker

'Print the Year_Change values
ws.Range("J" & Summary_Table_Row).Value = Year_Change

'Print the Year_Change values
ws.Range("K" & Summary_Table_Row).Value = Percentage_Change

'------------------------------------------------------
'MOVE DOWN ONE CELL TO AVOID OVERWRITE PREVIOUS ENTRY
'------------------------------------------------
'Add one to the summary table Row
Summary_Table_Row = Summary_Table_Row + 1


End If

Next i

Next ws

End Sub