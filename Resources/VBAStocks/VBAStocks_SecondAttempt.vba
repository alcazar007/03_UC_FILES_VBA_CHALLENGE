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
Dim Percent_Change As Double
Dim Total_Stock_Volume As Integer
Dim Summary_Table_Row As Integer
Dim Open_Date As Double
Dim Close_Date As Double



Summary_Table_Row = 2
Year_Change = 0
Percent_Change = 0
Total_Stock_Volume = 0

'-------------------------------------------------------------------------------
'ADD THE WORD TICKER, YEAR_CHANGE, PERCENTAGE_CHANGE, AND TOTAL_STOCK TO HEADERS
'-----------------------------------------------------------------------------

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Year_Change"
ws.Cells(1, 11).Value = "Percent_Change"
ws.Cells(1, 12).Value = "Total_Stock_Volume"


'this prevents my overflow error
On Error Resume Next
'---------------------------------------------------------
'FOR EACH WS IN WORKSHEETS
'------------------------------------------------------
WorksheetName = ws.Name
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
'------------------------------------------------------
'LOOP THROUGH ALL <TICKER> VALUES
'-------------------------------------------------------

For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

        '------------------------------------------------
        'PLACE PULLED VALUES TO SPECIFIC Cells I & L
        '-------------------------------------------------------

        'Print the Ticker Name
         ws.Range("I" & Summary_Table_Row).Value = Ticker
        'Print the Total_Stock_Volume values
         ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

        '------------------------------------------------------
        'MOVE DOWN ONE CELL TO AVOID OVERWRITE PREVIOUS ENTRY
        '------------------------------------------------
        Summary_Table_Row = Summary_Table_Row + 1

        '-------------------------------------------------
        'Reset Volume Total_Stock_Volume
        '--------------------------------------------------
        Total_Stock_Volume = 0

    Else

        '-----------------------------------------------------
        'Add To The Ticker Total 
        '-----------------------------------------------------
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    End If
        
    Next I

  End Sub



Sub Calculate_Changes()

'---------------------------------------------------------
'SET VARIABLES
'---------------------------------------------------------
Dim Open_Date As Long
Dim End_Row As Long
Dim Summary_Table_Row As Integer

Summary_Table_Row = 2
lastrow = lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'--------------------------------------------------------
'Loop through Ticker
'------------------------------------------------------
For i = 2 to lastrow

'Find Start and End for Rows
Start_Ticker = Range("A:A").Find(what:=Cells(i, 9), after:=Cells(1,1)).Row
End_Ticker = Range("A:A").Find(what:=Cells(i, 9), after:=Cells(1, 1),SearchDirection:=xlPrevious).Row

'------------------------------------------------------------
'Update Summary table
'------------------------------------------------------------
 'Print the Year_Change values
        ws.Range("K" & Summary_Table_Row).Value = Range("C" & Start_Ticker).Value - Range("F" & End_Ticker).Value

 


'--------------------------------------------------------
'START AND END ROWS WITH TICKER LOOP
'------------------------------------------------------

     '   Open_Date = ws.Cells(i, 3).Value
     '   Close_Date = ws.Cells(i, 6).Value
     '   Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
      '  Year_Change = Close_Date - Open_Date
      '  Percent_Change = (Close_Date - Open_Date) / Close_Date
      '  Percent_Change = Percent_Change * 100

       
'NEED TO PULL FIRST VALUE OF OPEN DATE AND LAST VALUE FOR CLOSED DATE
'------------------------------------------------



'------------------------------------------------
'PLACE PULLED VALUES TO SPECIFIC Cells
'-------------------------------------------------------
'Print the Ticker Name
      ' ws.Range("I" & Summary_Table_Row).Value = Ticker

        'Print the Year_Change values
       'ws.Range("J" & Summary_Table_Row).Value = Year_Change

     '   'Print the Year_Change values
     '   ws.Range("K" & Summary_Table_Row).Value = Percent_Change

     '   'Print the Total_Stock_Volume values
     '  ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
'


'--------------------------------------------------------
'FORMAT PERCENT COLUMN
'---------------------------------------------------------
'ws.Columns("K").NumberFormat = "0.00%"

'------------------------------------------------------
'MOVE DOWN ONE CELL TO AVOID OVERWRITE PREVIOUS ENTRY
'------------------------------------------------
'Add one to the summary table Row
'Summary_Table_Row = Summary_Table_Row + 1



End If

Next i

'MOVE ON TO NEXT SHEET
Next ws

End Sub