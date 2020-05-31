Attribute VB_Name = "Module1"
Sub StockData()

    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim TickerCount As Long
    Dim Ticker As String
    Dim TotalVol As Double
    Dim PercentChange As Double
     
        
  For Each ws In Worksheets
    
    TickerCount = 1
    TotalVol = 0
        
        For i = 2 To 705714
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                TotalVol = TotalVol + ws.Cells(i, 7).Value
            
                OpenValue = ws.Cells(i, 3).Value
            
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And OpenValue > 0 Then
            
                TotalVol = TotalVol + ws.Cells(i, 7).Value
                
                Ticker = ws.Cells(i, 1).Value
                
                TickerCount = TickerCount + 1
                
                ws.Cells(TickerCount, 9).Value = Ticker
                
                CloseValue = ws.Cells(i, 6).Value
                
                ws.Cells(TickerCount, 12).Value = TotalVol
                
                YearlyChange = CloseValue - OpenValue
                
                ws.Cells(TickerCount, 10).Value = YearlyChange
                
                ws.Cells(TickerCount, 11).Value = PercentChange
                
                PercentChange = (YearlyChange / OpenValue) * 100
                
                TotalVol = 0
                
                
            Else
                
                TotalVol = TotalVol + ws.Cells(i, 7).Value
                      
                
            End If
                
            Next i
                
            
    Next ws:


End Sub

