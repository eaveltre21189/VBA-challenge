Sub StockSummary()
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        Dim ticker As String
        
        Dim total_volume As Double
        total_volume = 0
        
        Dim first_open As Double
        Dim last_close As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Dim LR As Long
        LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To LR
            
            If ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                
                first_open = first_open
                
                last_close = ws.Cells(i, 6).Value
                
                yearly_change = last_close - first_open
                ws.Range("J" & Summary_Table_Row - 1).Value = yearly_change
                
                percent_change = yearly_change / first_open
                ws.Range("K" & Summary_Table_Row - 1).Value = percent_change
                
                total_volume = total_volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Table_Row - 1).Value = total_volume
            
            Else
                
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = ticker
                
                first_open = ws.Cells(i, 3).Value
                
                last_close = ws.Cells(i, 6).Value
                
                yearly_change = yearly_change
                
                percent_change = percent_change
                
                total_volume = 0
                
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                Summary_Table_Row = Summary_Table_Row + 1
            
            End If
        
        Next i
        
        Dim RngJ As Range
        Dim RngK As Range
        Dim ColorCell As Range        
        
        Set RngJ = ws.Range("J2", ws.Range("J2").End(xlDown))
        Set RngK = ws.Range("K2", ws.Range("K2").End(xlDown))             
        
        For Each ColorCell In RngJ
            If ColorCell.Value >= 0 Then
                ColorCell.Interior.Color = RGB(198, 239, 206)
            ElseIf ColorCell.Value < 0 Then
                ColorCell.Interior.Color = RGB(255, 199, 206)
            Else
                ColorCell.Interior.ColorIndex = xlNone
            End If
        Next
        
        For Each ColorCell In RngK
            If ColorCell.Value >= 0 Then
                ColorCell.Interior.Color = RGB(198, 239, 206)
            ElseIf ColorCell.Value < 0 Then
                ColorCell.Interior.Color = RGB(255, 199, 206)
            Else
                ColorCell.Interior.ColorIndex = xlNone
            End If
        Next
        
    Next ws

End Sub
