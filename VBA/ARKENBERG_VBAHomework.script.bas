Sub StockTrendsTest():

    'need to find the last row in  each sheet - they're huge
    'first part is basically the same as the Credit Card Solver exercise
    'but there are two discrete totals in the summary
        'sum TickerTotal and draw to Ticker cells
        'sum volumeTotal and draw to Volume cells
    'use worksheet range from Wells Fargo exercise to query all sheets in the workbook
    
    For Each ws In Worksheets
    
        'setting up the main variables
        Dim lastRow As String
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'ActiveWorkbook.Worksheets

        Dim ticker As String
        Dim tickerCount As Long

        Dim volume as Long
        Dim volumeCount as Long

        Dim rowSummary As Long
        rowSummary = 2


        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'set the ticker name and then add to the total
                ticker = ws.Cells(i, 1).Value
                tickerCount = tickerCount + 1

                volume = ws.Cells(i,7).Value
                volumeCount = volumeCount + 1
                'print the ticker to the summary, print the total volume to the summary
                ws.Range("J" & rowSummary).Value = ticker 
                ws.Range("k" & rowSummary).Value = volume
                'increment summary table and reset ticker total to zero
                rowSummary = rowSummary + 1
                tickerCount = 0

            Else: ticker = ticker + ws.Cells(i, 10).Value

            End If

        Next i

    Next ws

End Sub

