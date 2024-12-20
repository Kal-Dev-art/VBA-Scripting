
Sub CalculateStockDataWithTickerChange()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Paper(Q4)") ' Update to your actual sheet name
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' Set output headers
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Quarterly Change"
    ws.Cells(1, 12).Value = "Percentage Change"
    ws.Cells(1, 13).Value = "Total Volume"
    outputRow = 2 ' Start output from the second row
    totalVolume = 0 ' Initialize total volume
    ' Loop through the data
    For i = 2 To lastRow
        ' Get current ticker, opening price, closing price, and add to total volume
        ticker = ws.Cells(i, 1).Value
        openingPrice = ws.Cells(i, 3).Value
        closingPrice = ws.Cells(i, 6).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        ' Calculate quarterly change and percentage change
        quarterlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
            percentageChange = (quarterlyChange / openingPrice) * 100
        Else
            percentageChange = 0
        End If
        ' Check if the next row's ticker is different or we are at the last row
        If i = lastRow Or ws.Cells(i + 1, 1).Value <> ticker Then
            ' Output results for the current ticker
            ws.Cells(outputRow, 10).Value = ticker
            ws.Cells(outputRow, 11).Value = quarterlyChange
            ws.Cells(outputRow, 12).Value = percentageChange
            ws.Cells(outputRow, 13).Value = totalVolume
            ' Move to the next output row and reset total volume
            outputRow = outputRow + 1
            totalVolume = 0
        End If
    Next i
    MsgBox "Stock data calculation complete!"
End Sub
Sub AnalyzeStocks()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim stockWithGreatestIncrease As String
    Dim stockWithGreatestDecrease As String
    Dim stockWithGreatestVolume As String
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column A for the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' Initialize variables
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        ' Add headers for the output table in the current worksheet
        ws.Range("O1").Value = "Description"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ' Loop through the stocks to find the greatest values
        ' For i = 2 To lastRow
        '     ' Check for greatest percentage increase
        '     If ws.Cells(i, 2).Value > greatestIncrease Then
        '         greatestIncrease = ws.Cells(i, 2).Value
        '         stockWithGreatestIncrease = ws.Cells(i, 1).Value
        '     End If
        '     ' Check for greatest percentage decrease
        '     If ws.Cells(i, 2).Value < greatestDecrease Then
        '         greatestDecrease = ws.Cells(i, 2).Value
        '         stockWithGreatestDecrease = ws.Cells(i, 1).Value
        '     End If
        '     ' Check for greatest total volume
        '     If ws.Cells(i, 3).Value > greatestVolume Then
        '         greatestVolume = ws.Cells(i, 3).Value
        '         stockWithGreatestVolume = ws.Cells(i, 1).Value
        '     End If
        ' Next i




       greatestIncrease = WorksheetFunction.Max(ws.Range("L:L"))
       greatestDecrease = WorksheetFunction.Min(ws.Range("L:L"))
       greatestVolume = WorksheetFunction.Max(ws.Range("M:M"))
       matchpositionmaxpercent = WorksheetFunction.Match(greatestIncrease, ws.Range("L:L"), 0)
       stockWithGreatestIncrease = WorksheetFunction.Index(ws.Range("J:J"), matchpositionmaxpercent)
       matchpositionminpercent = WorksheetFunction.Match(greatestDecrease, ws.Range("L:L"), 0)
       stockWithGreatestDecrease = WorksheetFunction.Index(ws.Range("J:J"), matchpositionminpercent)
       matchpositiontotalvolume = WorksheetFunction.Match(greatestVolume, ws.Range("M:M"), 0)
       stockWithGreatestVolume = WorksheetFunction.Index(ws.Range("J:J"), matchpositiontotalvolume)
       
    

        ' Output the results to the worksheet
        ws.Range("P2").Value = stockWithGreatestIncrease
        ws.Range("Q2").Value = Format(greatestIncrease, "Percent")
        ws.Range("P3").Value = stockWithGreatestDecrease
        ws.Range("Q3").Value = Format(greatestDecrease, "Percent")
        ws.Range("P4").Value = stockWithGreatestVolume
        ws.Range("Q4").Value = Format(greatestVolume, "Scientific")
        ' Optional: Apply formatting for better readability
        ws.Columns("P:Q").AutoFit
NextWorksheet:
    Next ws
    MsgBox "Analysis complete for all worksheets!"
    
  End Sub

