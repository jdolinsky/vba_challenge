'returns a total count of rows in Sheet with data
Function GetRowCount()
    Dim lastRow As Long
    With ActiveSheet
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    GetRowCount = lastRow
End Function

Sub StockStats()
    Dim openPrice, closePrice, yearlyChange As Double
    Dim totalVolume As LongLong
    'Create summary headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    numRows = GetRowCount()
    startRow = 2
    summaryIndex = 2
    totalVolume = 0
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

