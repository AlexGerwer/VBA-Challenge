Sub ProcessStockData()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim i As Long
    Dim uniqueTickers As Variant
    Dim percentChange As Double
    Dim quarterlyChange As Double
    Dim totalVolume As LongLong
    Dim firstDataCell As Range 'Added to store first cell for conditional formatting
    Dim maxPercentChangeRow As Long, minPercentChangeRow As Long, maxVolumeRow As Long
    Dim maxPercentChange As Double, minPercentChange As Double, maxVolume As LongLong

    Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        With ws
            'Label columns (and autofit to header width initially)
            .Cells(1, "I").Value = "Ticker": .Columns("I").AutoFit
            .Cells(1, "J").Value = "Quarterly Change": .Columns("J").AutoFit
            .Cells(1, "K").Value = "Percent Change": .Columns("K").AutoFit
            .Cells(1, "L").Value = "Total Stock Volume": .Columns("L").AutoFit
            
            'Write headers for columns O, P and Q
            .Cells(1, "P").Value = "Ticker"
            .Cells(1, "Q").Value = "Value"
            .Cells(2, "O").Value = "Greatest % Increase"
            .Cells(3, "O").Value = "Greatest % Decrease"
            .Cells(4, "O").Value = "Greatest Total Volume"
            


            'Set last row to find the last row with data in column A
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

            If lastRow > 1 Then 'Check for data before proceeding
                'Get unique tickers from column A
                uniqueTickers = Application.WorksheetFunction.Unique(.Range("A2:A" & lastRow))

                'Keep track of the first data cell in column J for conditional formatting
                Set firstDataCell = .Cells(.Rows.Count, "J").End(xlUp).Offset(1)

                'Loop through each unique ticker
                For i = 1 To UBound(uniqueTickers, 1)
                    ticker = uniqueTickers(i, 1)

                    'Calculate values *before* writing to the sheet
                    quarterlyChange = .Cells(GetLatestDate(ws, ticker), "F").Value - .Cells(GetEarliestDate(ws, ticker), "C").Value
                    If .Cells(GetEarliestDate(ws, ticker), "C").Value = 0 Then
                        percentChange = 0
                    Else
                        percentChange = (.Cells(GetLatestDate(ws, ticker), "F").Value - .Cells(GetEarliestDate(ws, ticker), "C").Value) / .Cells(GetEarliestDate(ws, ticker), "C").Value
                    End If
                    totalVolume = WorksheetFunction.Sum(.Range(.Cells(GetEarliestDate(ws, ticker), "G"), .Cells(GetLatestDate(ws, ticker), "G")))

                    'Write the ticker and calculated values to the next available row
                    With .Cells(.Rows.Count, "I").End(xlUp).Offset(1)
                        .Value = ticker
                        .Offset(0, 1).Value = quarterlyChange
                        .Offset(0, 2).Value = percentChange
                        .Offset(0, 2).NumberFormat = "0.00%"
                        .Offset(0, 3).Value = totalVolume
                    End With
                Next i

                'Apply Conditional Formatting to Column J *after* data is written
                With .Range(firstDataCell, .Cells(.Rows.Count, "J").End(xlUp))
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    .FormatConditions(1).Interior.Color = vbRed
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                    .FormatConditions(2).Interior.Color = vbGreen
                End With


                 On Error Resume Next 'Handle potential errors if columns K or L are empty
                'Find the rows with max/min values (using Match is more efficient)
                maxPercentChangeRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(.Range("K:K")), .Range("K:K"), 0)
                maxPercentChange = .Cells(maxPercentChangeRow, "K").Value

                minPercentChangeRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(.Range("K:K")), .Range("K:K"), 0)
                minPercentChange = .Cells(minPercentChangeRow, "K").Value

                maxVolumeRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(.Range("L:L")), .Range("L:L"), 0)
                maxVolume = .Cells(maxVolumeRow, "L").Value

                On Error GoTo 0 'Reset error handling


                'Write the values and corresponding tickers to cells (only if rows were found)
                If maxPercentChangeRow > 0 Then
                    .Cells(2, "P").Value = .Cells(maxPercentChangeRow, "I").Value
                    .Cells(2, "Q").Value = maxPercentChange
                    .Cells(2, "Q").NumberFormat = "0.00%" 'format as percentage
                End If

                If minPercentChangeRow > 0 Then
                    .Cells(3, "P").Value = .Cells(minPercentChangeRow, "I").Value
                    .Cells(3, "Q").Value = minPercentChange
                    .Cells(3, "Q").NumberFormat = "0.00%" 'format as percentage
                End If

                 If maxVolumeRow > 0 Then
                    .Cells(4, "P").Value = .Cells(maxVolumeRow, "I").Value
                    .Cells(4, "Q").Value = maxVolume
                End If



            End If 'End check for data



            'Autofit columns *after* all data and formatting is applied, considering header width
            .Columns("I:Q").AutoFit  'Includes columns O, P, and Q now
            
           
            'Adjust column widths for headers if needed
            If .Cells(1, "I").ColumnWidth < Len(.Cells(1, "I").Value) + 1 Then .Columns("I").ColumnWidth = Len(.Cells(1, "I").Value) + 1
            If .Cells(1, "J").ColumnWidth < Len(.Cells(1, "J").Value) + 1 Then .Columns("J").ColumnWidth = Len(.Cells(1, "J").Value) + 1
             If .Cells(1, "K").ColumnWidth < Len(.Cells(1, "K").Value) + 1 Then .Columns("K").ColumnWidth = Len(.Cells(1, "K").Value) + 1
             If .Cells(1, "L").ColumnWidth < Len(.Cells(1, "L").Value) + 1 Then .Columns("L").ColumnWidth = Len(.Cells(1, "L").Value) + 1
              If .Cells(1, "O").ColumnWidth < Len(.Cells(1, "O").Value) + 1 Then .Columns("O").ColumnWidth = Len(.Cells(1, "O").Value) + 1
              If .Cells(1, "P").ColumnWidth < Len(.Cells(1, "P").Value) + 1 Then .Columns("P").ColumnWidth = Len(.Cells(1, "P").Value) + 1
              If .Cells(1, "Q").ColumnWidth < Len(.Cells(1, "Q").Value) + 1 Then .Columns("Q").ColumnWidth = Len(.Cells(1, "Q").Value) + 1

        End With
    Next ws
End Sub


' Helper function to get the row number of the earliest date for a given ticker
Function GetEarliestDate(ws As Worksheet, ticker As String) As Long
    Dim lastRow As Long, i As Long, earliestDate As Long, earliestRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    earliestDate = 2147483647 ' Maximum date value for Long, used as initial value

    For i = 2 To lastRow
        If ws.Cells(i, "A").Value = ticker Then
            If ws.Cells(i, "B").Value < earliestDate Then
                earliestDate = ws.Cells(i, "B").Value
                earliestRow = i
            End If
        End If
    Next i

    GetEarliestDate = earliestRow
End Function


' Helper function to get the row number of the latest date for a given ticker
Function GetLatestDate(ws As Worksheet, ticker As String) As Long
    Dim lastRow As Long, i As Long, latestDate As Long, latestRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    latestDate = 0  'used as initial value

    For i = 2 To lastRow
        If ws.Cells(i, "A").Value = ticker Then
            If ws.Cells(i, "B").Value > latestDate Then
                latestDate = ws.Cells(i, "B").Value
                latestRow = i
            End If
        End If
    Next i

    GetLatestDate = latestRow
End Function

