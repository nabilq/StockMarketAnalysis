Attribute VB_Name = "Module1"
Sub stock()
For Each ws In Worksheets

    'specify last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
     Columns("O:O").EntireColumn.AutoFit
    'specify percentage change as percentage
    ws.Range("k:k").NumberFormat = "0.00%"
    
    'Headers for main table
    
    ws.Range("I1:L1").Font.Bold = True
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    'Headers for secondary table
    
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"





  ' Set an initial variable for holding the ticker name


  ' Set an initial variable for holding the total per ticker volume
   ' Set an initial variable for holding the yearly change
   'set an an intial variable for holding the number of rows?
  Total = 0
  
  Yearly_change = 0

   Start = 0



  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock rows
  For i = 2 To lastrow
  
  
  


    ' Check if we are still within the same stock, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name
      ticker = ws.Cells(i, 1)
      
      counter = WorksheetFunction.CountIf(ws.Range("A:A"), ticker)
      
      counter2 = ws.Cells(i - (counter - 1), 3)
      
      If counter2 = 0 Then
      
      Else
      

      ' Add to the Ticker Volume Total
      Total = Total + ws.Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker

      ' Print the Ticker Volume Total to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total
      
  
      'Yearly Change formula
      Yearly_change = ws.Cells(i, 6) - ws.Cells(i - (counter - 1), 3)
      
      'print Yearly Change in Summary Table
      ws.Range("J" & Summary_Table_Row) = Yearly_change
  

      'Percent Change Formula


      percent_change = Yearly_change / counter2
    
      
      'Print Percent Change in Summary Table
      ws.Range("K" & Summary_Table_Row) = percent_change

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total
      Total = 0
     
      percent_change = 0
      counter2 = 0
 End If
    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Volume Total
      Total = Total + ws.Cells(i, 7).Value


End If


  Next i

 'conditional formatting for Yearly_change
   
    ws.Select
     ws.Range("J2:J290").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
   
  
    

 


'Format max and min percentage change as percentage
ws.Range("q2:q3").NumberFormat = "0.00%"


'Find Max percentage change and ticker and print to cell

        ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 16).Value = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match _
                        (ws.Cells(2, 17).Value, ws.Range("K:K"), 0))
                        
 'Find Min percentage change and  ticker and print to cell
 ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K:K"))
   ws.Cells(3, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match _
                        (ws.Cells(3, 17).Value, ws.Range("K:K"), 0))
                        

'Find Greatest Total Volume and Ticker

        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(4, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match _
                        (ws.Cells(4, 17).Value, ws.Range("L:L"), 0))
                        
Next ws

End Sub

