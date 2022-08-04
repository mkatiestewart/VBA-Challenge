Sub VBA_Home():

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim ticker As String
    Dim vol As Double
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim Percent_change As Double
    Dim Summary_Table_Row As Long

    Summary_Table_Row = 2
    
        For Each ws In ThisWorkbook.Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Changed"
        ws.Cells(1, 12).Value = "Total"
    
        year_open = ws.Cells(2, 3).Value
        For i = 2 To ws.UsedRange.Rows.Count
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                year_close = ws.Cells(i, 6).Value
                yearly_change = year_close - year_open
                
            End If
                
            If year_open <> 0 Then
                    Percent_change = (yearly_change / year_open) * 100
                
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = Percent_change
            ws.Cells(Summary_Table_Row, 12).Value = ticker
            Summary_Table_Row = Summary_Table_Row + 1
            
                vol = 0
                
            End If
            
            Next i
            
            Next ws
            
End Sub