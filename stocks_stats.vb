Sub Main()
    Dim wsCount As Integer
    Dim i As Integer
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    wsCount = ActiveWorkbook.Worksheets.Count
    ' Begin the loop.
    For i = 1 To wsCount
        Worksheets(i).Activate
       doStats
       getSummary
    Next i
    MsgBox ("Finished")
End Sub
Sub getSummary()
    'find the summary
    Dim i, k, numRows, greatestIncrease, greatestDecrease As Double
    Dim greatestVolume As LongLong
    Dim tickerMax, tickerMin As String
    numRows = Cells(Rows.Count, "J").End(xlUp).Row
    greatestIncrease = Range("K2").Value
    greatestDecrease = Range("K2").Value
    tickerMax = Range("I2").Value
    tickerMin = Range("I2").Value
    tickerMaxVolume = Range("I2").Value
    greatestVolume = Range("L2").Value
    For i = 2 To numRows
        If greatestIncrease < Range("K" & i + 1).Value Then
            greatestIncrease = Range("K" & i + 1).Value
            tickerMax = Range("I" & i + 1).Value
        End If
        If greatestDecrease > Range("K" & i + 1).Value Then
            greatestDecrease = Range("K" & i + 1).Value
            tickerMin = Range("I" & i + 1).Value
        End If
        If greatestVolume < Range("L" & i + 1).Value Then
            greatestVolume = Range("L" & i + 1).Value
            tickerMaxVolume = Range("I" & i + 1).Value
        End If
    Next i
    'write summary
    'create headers
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    'format cells
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Range("P2").Value = tickerMax
    Range("P3").Value = tickerMin
    Range("P4").Value = tickerMaxVolume
    Range("Q2").Value = greatestIncrease
    Range("Q3").Value = greatestDecrease
    Range("Q4").Value = greatestVolume
End Sub

Sub doStats()
    Dim openPrice, closePrice, yearlyChange, greatestIncrease As Double
    Dim totalVolume As LongLong
    'Create summary headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    numRows = Cells(Rows.Count, "A").End(xlUp).Row
    startRow = 2
    summaryIndex = 2
    totalVolume = greatestIncrease = 0
    'set first open price for the first ticker in Sheet
    openPrice = Cells(startRow, 3).Value
    For i = startRow To numRows
        'check if the current ticker equals the next one
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'since we are on the last day of the year for the same stock, grab
        'the Price Closed on the last day
            closePrice = Cells(i, 6).Value
            'calculate yearly change
            yearlyChange = closePrice - openPrice
            'calculate percent change
            percentChange = yearlyChange / openPrice
            ' add last day volume
            totalVolume = totalVolume + Cells(i, 7).Value
            'write the summary and clean up
            Range("I" & CStr(summaryIndex)) = Cells(i, 1).Value
            Range("J" & CStr(summaryIndex)) = yearlyChange
            With Range("J" & CStr(summaryIndex))
                If yearlyChange <= 0 Then
                    .Interior.Color = RGB(250, 0, 0)
                Else
                    .Interior.Color = RGB(0, 250, 0)
                End If
            End With
            Range("K" & CStr(summaryIndex)) = percentChange
            Range("K" & CStr(summaryIndex)).NumberFormat = "0.00%"
            Range("L" & CStr(summaryIndex)) = totalVolume
            'store first Price Open in the beginning of the year
            openPrice = Cells(i + 1, 3).Value
            summaryIndex = summaryIndex + 1
            totalVolume = 0
        Else
            totalVolume = totalVolume + Cells(i, 7).Value
        End If
    Next i
    
End Sub



