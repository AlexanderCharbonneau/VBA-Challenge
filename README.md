# VBA-Challenge

I just added the macro workbook so it was all in one place
(I realized it did not save correctly so here is the vba if you want to copy paste it to see it work, sorry grader

"Sub StockAnalysis()
    
    ' Declarations of variables
    Dim initialPrice As Double
    Dim finalPrice As Double
    Dim stockSymbol As String
    Dim currentSymbol As String
    Dim totalTradingVolume As LongLong
    Dim currentSheet As Worksheet
    
    Dim resultRow As Long
    Dim dataRow As Long
    
    ' Loop through each worksheet in the workbook
    For Each currentSheet In ThisWorkbook.Worksheets
        
        ' Initialize variables
        initialPrice = 0
        finalPrice = 0
        stockSymbol = ""
        currentSymbol = ""
        resultRow = 1
        
        ' Loop through rows of data
        For dataRow = 2 To 760000
            totalTradingVolume = 0
            initialPrice = currentSheet.Cells(dataRow, 3).Value ' Start on C2
            currentSymbol = currentSheet.Cells(dataRow, 1).Value ' Start on A2
            stockSymbol = currentSymbol
            
            ' Exit loop if stockSymbol is empty
            If stockSymbol = "" Then
                Exit For
            End If
            
            resultRow = resultRow + 1
            
            ' Loop to calculate total trading volume and find final price
            Do While currentSymbol = stockSymbol
                totalTradingVolume = totalTradingVolume + currentSheet.Cells(dataRow, 7).Value
                nextSymbol = currentSheet.Cells(dataRow + 1, 1).Value
                
                ' Exit loop if the next symbol is different
                If Not currentSymbol = nextSymbol Then
                    finalPrice = currentSheet.Cells(dataRow, 6).Value
                    Exit Do
                End If
                
                dataRow = dataRow + 1
                currentSymbol = nextSymbol
            Loop
            
            ' Display outputs
            currentSheet.Cells(resultRow, 9).Value = stockSymbol
            currentSheet.Cells(resultRow, 10).Value = finalPrice - initialPrice
            currentSheet.Cells(resultRow, 11).Value = ((finalPrice - initialPrice) / initialPrice)
            currentSheet.Cells(resultRow, 12).Value = totalTradingVolume
        Next dataRow
    Next currentSheet
End Sub

Sub FindGreatest()
    
    ' Declarations
    Dim dataRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As LongLong
    Dim currentChange As Double
    Dim currentVolume As LongLong
    Dim stockSymbol As String
    Dim currentSheet As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each currentSheet In ThisWorkbook.Worksheets
        
        ' Initialize variables
        maxIncrease = 0
        maxDecrease = 0
        stockSymbol = ""
        
        ' Find the maximum percentage increase
        For dataRow = 2 To 3001
            currentChange = currentSheet.Cells(dataRow, 11).Value
            If currentChange > maxIncrease Then
                maxIncrease = currentChange
                stockSymbol = currentSheet.Cells(dataRow, 9).Value
            End If
        Next dataRow
        
        currentSheet.Cells(2, 16).Value = stockSymbol
        currentSheet.Cells(2, 17).Value = maxIncrease
        
        ' Find the maximum percentage decrease
        stockSymbol = ""
        For dataRow = 2 To 3001
            currentChange = currentSheet.Cells(dataRow, 11).Value
            If currentChange < maxDecrease Then
                maxDecrease = currentChange
                stockSymbol = currentSheet.Cells(dataRow, 9).Value
            End If
        Next dataRow
        
        currentSheet.Cells(3, 16).Value = stockSymbol
        currentSheet.Cells(3, 17).Value = maxDecrease
        
        ' Find the maximum total trading volume
        stockSymbol = ""
        maxVolume = 0
        For dataRow = 2 To 3001
            currentVolume = currentSheet.Cells(dataRow, 12).Value
            If currentVolume > maxVolume Then
                maxVolume = currentVolume
                stockSymbol = currentSheet.Cells(dataRow, 9).Value
            End If
        Next dataRow
        
        currentSheet.Cells(4, 16).Value = stockSymbol
        currentSheet.Cells(4, 17).Value = maxVolume
    Next currentSheet
End Sub
")
