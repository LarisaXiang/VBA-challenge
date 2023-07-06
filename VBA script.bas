Attribute VB_Name = "Module1"
Sub stockData()

    For Each ws In Worksheets
        
        Dim i As Long
        Dim j As Long
        Dim tickerCount As Long
        Dim lastRow_A As Long
            lastRow_A = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        Dim Percentage As Double
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        tickerCount = 2
        j = 2
        
'     For i = 2 To lastRow_A
'
'        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
'        ws.Cells(tickerCount, 9).Value = ws.Cells(i, 1).Value
'        ws.Cells(tickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
'        range(ws.Cells(tickerCount, 10).Address).NumberFormat = "0.00"
'            If ws.Cells(tickerCount, 10).Value < 0 Then
'            ws.Cells(tickerCount, 10).Interior.ColorIndex = 3
'
'            Else
'
'             ws.Cells(tickerCount, 10).Interior.ColorIndex = 4
'
'             End If
'        If ws.Cells(j, 3).Value <> 0 Then
'        Percentage = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
'        ws.Cells(tickerCount, 11).Value = Format(Percentage, "Percent")
'
'        Else
'        ws.Cells(tickerCount, 11).Value = Format(0, "Percent")
'
'        End If
'
'        ws.Cells(tickerCount, 12) = WorksheetFunction.Sum(range(ws.Cells(j, 7), ws.Cells(i, 7)))
'
'        tickerCount = tickerCount + 1
'
'
'        j = i + 1
'
'End If
'       Next i


    Dim lastRow_I As Long
            lastRow_I = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    
    Dim rng As range
    Dim rng1 As range
    Set rng = range("K1:K" & lastRow_I)
    Set rng1 = range("L1:L" & lastRow_I)
    Dim GreatIncrTicker, GreatDecrTicker, GreatVolValTicker As Integer
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim GreatIncrVal
    Dim GreatDecrVal
    Dim GreatVolVal

        
        
        GreatIncr = WorksheetFunction.Max(rng)
        ws.Cells(2, 17).Value = GreatIncr
        ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
        
        GreatDecr = WorksheetFunction.Min(rng)
        ws.Cells(3, 17).Value = GreatDecr
        ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
               
        GreatIncrVal = Format(GreatIncr, "Percent")
        
        
        GreatDecrVal = Format(GreatDecr, "Percent")
               
        GreatIncrTicker = rng.Find(What:=GreatIncrVal, LookIn:=xlValues, LookAt:=xlWhole).row
        ws.Cells(2, 16).Value = Cells(GreatIncrTicker, 9).Value
        
        GreatDecrTicker = rng.Find(What:=GreatDecrVal, LookIn:=xlValues, LookAt:=xlWhole).row
        ws.Cells(3, 16).Value = Cells(GreatDecrTicker, 9).Value
        
    Dim GreatVol As Double
        
        GreatVol = WorksheetFunction.Max(rng1)
        
        ws.Cells(4, 17).Value = GreatVol
        
        GreatVolVal = Format(ws.Cells(4, 17).Value, "0.00000E+0")
        
        ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
    
        GreatVolValTicker = rng1.Find(What:=GreatVolVal, LookIn:=xlValues, LookAt:=xlWhole).row
        ws.Cells(4, 16).Value = Cells(GreatVolValTicker, 9).Value
        
        
        Next ws
        
    End Sub
        
        
        
        
