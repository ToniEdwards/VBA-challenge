Sub WallstreetFinal()
'Assign Variables
Dim i As Long
Dim ticker As String
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalStockVolume As Double
Dim opn As Double
Dim cls As Double
Dim lastRow As Long
'Loop through worksheets
For Each ws In Worksheets
    lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
'Loop through rows to populate cells with data
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        opn = ws.Cells(i, 3).Value
        cls = ws.Cells(i, 6).Value
        ws.Cells(i, 9).Value = ticker
        yearlyChange = cls - opn
        ws.Cells(i, 10).Value = yearlyChange
'Highlight positive vs negative Change
        If yearlyChange > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            percentChange = yearlyChange / opn
            ws.Cells(i, 11).Value = percentChange
            ws.Cells(i, 11).NumberFormat = "0.00%"
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
            ws.Cells(i, 11).Value = "0.00%"
        End If
'Check if the previous ticker and current ticker match
        'If they match increase and return the total stock volume
        If ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value Then
            totalStockVolume = ws.Cells(i - 1, 12).Value + ws.Cells(i, 7).Value
            ws.Cells(i, 12).Value = totalStockVolume
        'If they do not match return the total stock volume
        Else
            totalStockVolume = ws.Cells(i, 7).Value
            ws.Cells(i, 12).Value = totalStockVolume
        End If
    Next i
Next ws
End Sub

