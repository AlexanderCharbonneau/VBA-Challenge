Sub StockCounter()
    Dim sheet As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim tickerName As String
    Dim openAmount As Double
    Dim closeAmount As Double
    Dim volAmount As Double
    Dim isFirst As Boolean
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim outputRow As Long
    Dim bestInc As Double
    Dim bestDec As Double
    Dim bestVol As Double
    Dim bestIncName As String
    Dim bestDecName As String
    Dim bestVolName As String

    For Each sheet In ThisWorkbook.Sheets
        lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
        outputRow = 2 ' Initialize output row

        ' Clear existing formatting and data
        sheet.Cells.ClearFormats
        sheet.Range("I2:Q" & lastRow).ClearContents

        ' Set header values
        sheet.Cells(1, 9).Value = "Ticker"
        sheet.Cells(1, 10).Value = "Yearly Change"
        sheet.Cells(1, 11).Value = "Percent Change"
        sheet.Cells(1, 12).Value = "Total Stock Volume"
        sheet.Cells(1, 16).Value = "Ticker"
        sheet.Cells(1, 17).Value = "Value"
        sheet.Cells(2, 15).Value = "Greatest % Increase"
        sheet.Cells(3, 15).Value = "Greatest % Decrease"
        sheet.Cells(4, 15).Value = "Greatest Total Volume"

        isFirst = True ' Initialize isFirst

        For rowCount = 2 To lastRow
            If isFirst Then
                tickerName = sheet.Cells(rowCount, 1).Value
                openAmount = sheet.Cells(rowCount, 3).Value
                isFirst = False
            End If

            volAmount = volAmount + sheet.Cells(rowCount, 7).Value

            If sheet.Cells(rowCount + 1, 1).Value <> tickerName Then
                closeAmount = sheet.Cells(rowCount, 6).Value
                yearlyChange = closeAmount - openAmount
                If openAmount <> 0 Then
                    percentChange = yearlyChange / openAmount
                Else
                    percentChange = 0
                End If

                ' Output values
                outputRow = outputRow + 1
                sheet.Cells(outputRow, 9).Value = tickerName
                sheet.Cells(outputRow, 10).Value = yearlyChange
                sheet.Cells(outputRow, 11).Value = percentChange
                sheet.Cells(outputRow, 11).NumberFormat = "0.00%"
                sheet.Cells(outputRow, 12).Value = volAmount

                ' Apply color conditional formatting for the Yearly Change
                If yearlyChange > 0 Then
                    sheet.Cells(outputRow, 10).Interior.ColorIndex = 4 ' Green
                ElseIf yearlyChange < 0 Then
                    sheet.Cells(outputRow, 10).Interior.ColorIndex = 3 ' Red
                End If

                ' Check for greatest values
                If volAmount > bestVol Then
                    bestVol = volAmount
                    bestVolName = tickerName
                End If

                If percentChange > bestInc Then
                    bestInc = percentChange
                    bestIncName = tickerName
                ElseIf percentChange < bestDec Then
                    bestDec = percentChange
                    bestDecName = tickerName
                End If

                ' Reset variables for the next ticker
                volAmount = 0
                isFirst = True
            End If
        Next rowCount

        ' Output greatest values
        sheet.Cells(2, 16).Value = bestIncName
        sheet.Cells(2, 17).NumberFormat = "0.00%"
        sheet.Cells(2, 17).Value = bestInc
        sheet.Cells(3, 16).Value = bestDecName
        sheet.Cells(3, 17).NumberFormat = "0.00%"
        sheet.Cells(3, 17).Value = bestDec
        sheet.Cells(4, 16).Value = bestVolName
        sheet.Cells(4, 17).Value = bestVol
        sheet.Columns("O:O").AutoFit
        sheet.Columns("Q:Q").ColumnWidth = 11.3
        sheet.Range("I1:P1").Columns.AutoFit
    Next sheet
End Sub

