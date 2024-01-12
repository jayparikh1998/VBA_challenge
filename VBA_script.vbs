Attribute VB_Name = "Module1"
Sub AnalyzeStockDataInAllSheets()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
        AnalyzeStockData ws
    Next ws
End Sub


Sub AnalyzeStockData(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim startRow As Integer
    startRow = 2 ' Assuming row 1 has headers

    ' Variables to track greatest % increase, decrease, and total volume
    Dim greatestPctIncrease As Double
    Dim greatestPctDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestPctIncrease As String
    Dim tickerGreatestPctDecrease As String
    Dim tickerGreatestVolume As String
    

    greatestPctIncrease = 0
    greatestPctDecrease = 0
    greatestVolume = 0

    ' Setting up headers for output
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Volume"

    Dim outputRow As Integer
    outputRow = 2 ' Start row for output

    For i = startRow To lastRow
        If ws.Cells(i, 1).Value <> ticker And ticker <> "" Then
            ' Calculate and output data
            closingPrice = ws.Cells(i - 1, 6).Value
            pctChange = 0
            If openingPrice <> 0 Then
                pctChange = (closingPrice - openingPrice) / openingPrice
            End If

            ' Check for greatest % increase and decrease
            If pctChange > greatestPctIncrease Then
                greatestPctIncrease = pctChange
                tickerGreatestPctIncrease = ticker
            ElseIf pctChange < greatestPctDecrease Then
                greatestPctDecrease = pctChange
                tickerGreatestPctDecrease = ticker
            End If

            ' Check for greatest total volume
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                tickerGreatestVolume = ticker
            End If
            
            
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = closingPrice - openingPrice
            ws.Cells(outputRow, 11).Value = pctChange
            ws.Cells(outputRow, 12).Value = totalVolume
            
            ' Reset for next ticker
            openingPrice = ws.Cells(i, 3).Value
            totalVolume = 0
            outputRow = outputRow + 1
        End If

        If ticker = "" Then
            openingPrice = ws.Cells(i, 3).Value ' Open price of the stock
        End If

        ticker = ws.Cells(i, 1).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    Next i


    ' Output the last ticker
    closingPrice = ws.Cells(lastRow, 6).Value
    Dim lastPctChange As Double
    lastPctChange = 0
    If openingPrice <> 0 Then
        lastPctChange = (closingPrice - openingPrice) / openingPrice
    End If

    ' Check for greatest % increase and decrease for the last ticker
    If lastPctChange > greatestPctIncrease Then
        greatestPctIncrease = lastPctChange
        tickerGreatestPctIncrease = ticker
    ElseIf lastPctChange < greatestPctDecrease Then
        greatestPctDecrease = lastPctChange
        tickerGreatestPctDecrease = ticker
    End If

    ' Check for greatest total volume for the last ticker
    If totalVolume > greatestVolume Then
        greatestVolume = totalVolume
        tickerGreatestVolume = ticker
    End If

    ws.Cells(outputRow, 9).Value = ticker
    ws.Cells(outputRow, 10).Value = closingPrice - openingPrice
    ws.Cells(outputRow, 11).Value = lastPctChange
    ws.Cells(outputRow, 12).Value = totalVolume

    ' Output greatest % increase, decrease, and total volume
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = tickerGreatestPctIncrease
    ws.Cells(3, 15).Value = greatestPctIncrease

    ws.Cells(4, 14).Value = "Greatest % Decrease"
    ws.Cells(5, 14).Value = tickerGreatestPctDecrease
    ws.Cells(5, 15).Value = greatestPctDecrease

    ws.Cells(6, 14).Value = "Greatest Total Volume"
    ws.Cells(7, 14).Value = tickerGreatestVolume
    ws.Cells(7, 15).Value = greatestVolume
End Sub

