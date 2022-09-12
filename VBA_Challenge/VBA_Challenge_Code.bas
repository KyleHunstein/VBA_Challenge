Attribute VB_Name = "Module1"
Sub Button1_Click()

Range("I:L").EntireColumn.Insert
Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"

Dim ticker As String
Dim ticker_total As Double
ticker_total = 0

Dim info_table_row As Integer
info_table_row = 2

Dim open_value As Double
open_value = Cells(2, 3).Value
Dim close_value As Double

Dim condition1 As FormatCondition
Dim condition2 As FormatCondition

Dim Last_Row As Long
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To Last_Row
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        ticker = Cells(i, 1).Value
    
        ticker_total = ticker_total + Cells(i, 7).Value
        
        close_value = Cells(i, 6).Value
        
        Range("J" & info_table_row).Value = (close_value - open_value)
            
              '  Set condition1 = Range("J" & info_table_row).FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
               ' Set condition2 = Range("J" & info_table_row).FormatConditions.Add(xlCellValue, xlLess, "=0")
        
               ' With FormatConditions(1)
                
                    
                 '   Interior.ColorIndex = 3
                 '   'RGB(0, 255, 0)
               ' End With
              '  With condition2
                '    .ColorIndex = 3
                    'Interior.Color = RGB(255, 0, 0)
              '  End With
            If Range("J" & info_table_row).Value >= 0 Then
                Range("J" & info_table_row).Interior.Color = RGB(0, 255, 0)
            ElseIf Range("J" & info_table_row).Value < 0 Then
                Range("J" & info_table_row).Interior.Color = RGB(255, 0, 0)
            End If
    
        Range("K" & info_table_row).Value = FormatPercent(((close_value - open_value) / open_value))
    
        Range("I" & info_table_row).Value = ticker
    
        Range("L" & info_table_row).Value = ticker_total
    
        info_table_row = info_table_row + 1
    
        ticker_total = 0
        
        open_value = Cells(i + 1, 3).Value
    
    Else
        ticker_total = ticker_total + Cells(i, 7).Value
    End If
Next i
  

End Sub
