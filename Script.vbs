
Sub ticker()

Dim ticker As String
Dim yearly_change As Double
Dim total_stock_volume As Double
Dim open_price As Double
Dim close_price As Double
Dim first_day As Double
Dim yearly_from_zero As Double

yearly_change = 0
total_stock_volume = 0
open_price = 0
close_price = 0
yearly_from_zero = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

    first_day = 2

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ticker = Cells(i, 1).Value
      
      open_price = Cells(first_day, 3).Value
      
      first_day = i + 1
      
      close_price = Cells(i, 6).Value

      total_stock_volume = total_stock_volume + Cells(i, 7).Value

      ' Printing in the Summary Table
      
      Range("I" & Summary_Table_Row).Value = ticker
   
      Range("M" & Summary_Table_Row).Value = open_price

      Range("N" & Summary_Table_Row).Value = close_price
      
      Range("L" & Summary_Table_Row).Value = total_stock_volume
      
      open_price = Cells(first_day, 3).Value
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      total_stock_volume = 0

    Else
      
      total_stock_volume = total_stock_volume + Cells(i, 7).Value
      
    End If

  Next i

    lastrow_summary = Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow_summary
    
    yearly_change = Cells(i, 14).Value - Cells(i, 13).Value
    Cells(i, 10) = yearly_change
    
    If Cells(i, 13).Value = 0 Then

    Cells(i, 11).Value = 0
    
    Else
    
    Cells(i, 11).Value = (Cells(i, 10).Value * 100) / Cells(i, 13)
    Cells(i, 11).Value = WorksheetFunction.Round(Cells(i, 11).Value, 2)
        
    If Cells(i, 10) = 0 Then Cells(i, 11) = 0
    If Cells(i, 10).Value > 0 Then Cells(i, 10).Interior.ColorIndex = 4
    If Cells(i, 10).Value < 0 Then Cells(i, 10).Interior.ColorIndex = 3
    
    Cells(2, 18).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary))
    Cells(3, 18).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary))
    Cells(4, 18).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary))
    
    
    End If

 Next i

End Sub

Public Sub Reset()

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

    lastrow_summary = Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastrow_summary
       
    Cells(i, 9).Value = Empty
    Cells(i, 10).Value = Empty
    Cells(i, 11).Value = Empty
    Cells(i, 10).Interior.ColorIndex = x1None
    Cells(i, 12).Value = Empty
    Cells(i, 13).Value = Empty
    Cells(i, 14).Value = Empty
    Cells(i, 17).Value = Empty
    Cells(i, 18).Value = Empty

    Next i
    
End Sub