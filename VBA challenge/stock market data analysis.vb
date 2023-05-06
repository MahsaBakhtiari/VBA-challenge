Sub Yearly_Analysis()
    Dim LR As Long
    Dim newrow As Integer
    Dim currentcc As String
    Dim First_opening As Double
    Dim Last_closing As Double
    Dim ws As Worksheet
    
    
    
    For Each ws In ThisWorkbook.Worksheets
        LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
        newrow = 1
        ws.Cells(2, 18).Value = 0
        ws.Cells(3, 18).Value = 1
        ws.Cells(4, 18).Value = 0
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"
        ws.Cells(2, 16).Value = "Greatest % increase"
        ws.Cells(3, 16).Value = "Greatest % decrease"
        ws.Cells(4, 16).Value = "Greatest total volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Columns("L").NumberFormat = "$#,##0.00"
        ws.Columns("M").NumberFormat = "0.00%"
        ws.Range("R2:R3").NumberFormat = "0.00%"
        
        For i = 2 To LR
            If (ws.Cells(i, 1).Value) = ws.Cells(newrow, 11).Value Then
                ws.Cells(newrow, 14).Value = ws.Cells(newrow, 14).Value + ws.Cells(i, 7).Value
            Else
                newrow = newrow + 1
                ws.Cells(newrow, 11).Value = ws.Cells(i, 1).Value
                ws.Cells(newrow, 14).Value = ws.Cells(i, 7).Value
                If newrow > 2 Then
                    Last_closing = ws.Cells(i - 1, 6).Value
                    ws.Cells(newrow - 1, 12).Value = Last_closing - First_opening
                    If ws.Cells(newrow - 1, 12).Value < 0 Then
                        ws.Cells(newrow - 1, 12).Interior.ColorIndex = 3
                    Else
                        ws.Cells(newrow - 1, 12).Interior.ColorIndex = 4
                    End If
                    ws.Cells(newrow - 1, 13).Value = ws.Cells(newrow - 1, 12) / First_opening
                    If ws.Cells(newrow - 1, 13).Value > ws.Cells(2, 18).Value Then
                        ws.Cells(2, 18).Value = ws.Cells(newrow - 1, 13).Value
                        ws.Cells(2, 17).Value = ws.Cells(newrow - 1, 11).Value
                    End If
                    If ws.Cells(newrow - 1, 13).Value < ws.Cells(3, 18).Value Then
                        ws.Cells(3, 18).Value = ws.Cells(newrow - 1, 13).Value
                        ws.Cells(3, 17).Value = ws.Cells(newrow - 1, 11).Value
                    End If
                    If ws.Cells(newrow - 1, 14).Value > ws.Cells(4, 18).Value Then
                        ws.Cells(4, 18).Value = ws.Cells(newrow - 1, 14).Value
                        ws.Cells(4, 17).Value = ws.Cells(newrow - 1, 11).Value
                    End If
                End If
                First_opening = ws.Cells(i, 3).Value
            End If
        Next i
    Next ws
End Sub
