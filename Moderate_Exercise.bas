Attribute VB_Name = "Module2"
Sub Ticker_Yearly_Change()

  ' Set an initial variable for holding the Ticker Symbol and Volume
  Dim Ticker_Symbol As String
  Dim Stock_Vol_Total As Double
  Stock_Vol_Total = 0
  
  ' Set an initial variable for holding the Opening Price  per Ticker
  Dim Opening_Price As Double
  Dim Closing_Price As Double
  Dim Yearly_difference As Double
  'Opening_Price = 0
  'Closing_Price = 0
  Yearly_difference = 0

  ' Keep track of the location for each Ticker Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Clear the summary table
  Columns("I:Q").Select
  Selection.Clear

'Set the heading for the summary table
Range("I1").Value = "Ticker_Symbol"
Range("J1").Value = "Yearly_Change"
Range("K1").Value = "Percent_Change"
Range("L1").Value = "Total_Stock_Volume"

Dim i As Double
Dim x As Double

x = 2
i = 2

 Cells(x, 9).Value = Cells(x, 1).Value
 Opening_Price = Cells(i, 3).Value
 
  ' Loop through all Ticker Symbol Volume
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To LastRow

    ' Check if we are still within the same Ticker Symbol, if it is not...
    'If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

   
    If Cells(i, 1).Value = Cells(x, 9).Value Then

      ' Set the Ticker Symbol
      Ticker_Symbol = Cells(i, 1).Value

      ' Add to the Ticker Volume Total
      Stock_Vol_Total = Stock_Vol_Total + Cells(i, 7).Value

      ' Find the Opening Price
      'Opening_Price = Cells(i, 3).Value
          
      ' Find the last Closing Price
      Closing_Price = Cells(i, 6).Value
      
      Else
      
         
       Cells(x, 10).Value = Closing_Price - Opening_Price
       
        If Closing_Price <= 0 Then
        
       Cells(x, 11).Value = 0
                    
                    Else
                    
                    Cells(x, 11).Value = (Closing_Price - Opening_Price) / Closing_Price
                
                End If
                
                Cells(x, 11).Style = "Percent"
                
                
            If Cells(x, 10).Value >= 0 Then
                
                Cells(x, 10).Interior.ColorIndex = 4
                 
                  Else
                  
                Cells(x, 10).Interior.ColorIndex = 3
                
            End If
            
        Cells(x, 12).Value = Stock_Vol_Total
        
   
    x = x + 1
    Cells(x, 9).Value = Cells(i, 1).Value
    
    End If
    Opening_Price = Cells(i, 3).Value
    Next i
    
    
    Cells(x, 10).Value = Closing_Price - Opening_Price
        If Closing_Price <= 0 Then
            Cells(x, 11).Value = 0
            
            Else
            
            Cells(x, 11).Value = (Closing_Price - Opening_Price) / Opening_Price
            
        End If
        
            Cells(x, 11).Style = "Percent"
            
            If Cells(x, 10).Value >= 0 Then
                Cells(x, 10).Interior.ColorIndex = 4
                    Else
                Cells(1, 10).Interior.ColorIndex = 3
                
            End If
            
    Cells(x, 12).Value = Stock_Vol_Total
    
Cells(1, 1).Select

End Sub


