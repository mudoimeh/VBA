Attribute VB_Name = "Module1"
Sub SheetLoop()
For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        Sheets(ws.Name).Select
        Call Ticker_Stock_Volume
    Next ws
    
End Sub


Sub Ticker_Stock_Volume()

  ' Set an initial variable for holding the Ticker Symbol
  Dim Ticker_Symbol As String

  ' Set an initial variable for holding the total Stock Volume per Ticker
  Dim Stock_Vol_Total As Double
  Stock_Vol_Total = 0

  ' Keep track of the location for each Ticker Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Clear the summary table
  Columns("I:Q").Select
  Selection.Clear
  
'Set the heading for the summary table
Range("I1").Value = "Ticker_Symbol"
Range("J1").Value = "Total_Stock_Volume"


  ' Loop through all Ticker Symbol Volume
  
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To LastRow

    ' Check if we are still within the same Ticker Symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Symbol
      Ticker_Symbol = Cells(i, 1).Value

      ' Add to the Ticker Volume Total
      Stock_Vol_Total = Stock_Vol_Total + Cells(i, 7).Value

      ' Print the Ticker Symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Symbol

      ' Print the Stock Volume to the Summary Table
      Range("J" & Summary_Table_Row).Value = Stock_Vol_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume Total
      Stock_Vol_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Stock Volume Total
      Stock_Vol_Total = Stock_Vol_Total + Cells(i, 7).Value

    End If

  Next i
  
End Sub



