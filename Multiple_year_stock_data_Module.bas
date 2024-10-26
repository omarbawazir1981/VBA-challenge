Attribute VB_Name = "Module1"
Sub StockAnalysisWithoutDoWhile()

    ' Declare variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim change As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim lastRow As Long
    Dim i As Long
    Dim summaryRow As Long
    Dim firstRowForTicker As Long
    Dim firstSheet As Worksheet
    
    ' Reference to the first sheet (for consolidated results)
    Set firstSheet = ThisWorkbook.Worksheets(1)
    
    ' Initialize the greatest values
    greatestIncrease = -999999
    greatestDecrease = 999999
    greatestVolume = 0
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row in column A (ticker column)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2 ' Start the summary row from row 2
        
        ' Add headers in columns I, J, K, and L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through each row of stock data
        For i = 2 To lastRow
            
            ' Check if this is the first appearance of a new ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' New ticker detected, set the opening price and reset total volume
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
                firstRowForTicker = i
            End If
            
            ' Add the volume for the current row
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if this is the last row for the current ticker
            If i = lastRow Or ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' This is the last row for this ticker, capture the closing price
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate the change and percentage change
                change = closePrice - openPrice
                If openPrice <> 0 Then
                    percentageChange = (closePrice - openPrice) / openPrice ' Percent change
                Else
                    percentageChange = 0
                End If
                
                ' Output the results for this ticker in the summary table
                ws.Cells(summaryRow, 9).Value = ticker ' Ticker symbol
                ws.Cells(summaryRow, 10).Value = change ' Price change
                ws.Cells(summaryRow, 11).Value = percentageChange ' Percentage change
                ws.Cells(summaryRow, 12).Value = totalVolume ' Total volume
                
                ' Apply conditional formatting (color)
                If change > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = vbGreen ' Green for positive change
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = vbRed ' Red for negative change
                End If
                
                ' Check if this ticker has the greatest increase, decrease, or total volume
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    greatestIncreaseTicker = ticker
                End If
                
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    greatestDecreaseTicker = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
                
                ' Move to the next row in the summary table
                summaryRow = summaryRow + 1
            End If
            
        Next i
        
        ' Sort the summary data by ticker
        ws.Range("I2:L" & summaryRow - 1).Sort Key1:=ws.Range("I2"), Order1:=xlAscending
        
        ' AutoFit the columns for the output
        ws.Columns("I:L").AutoFit
        
        ' Format Percent Change and Total Volume columns
        ws.Range("K2:K" & summaryRow - 1).NumberFormat = "0.00%" ' Percent change format
        ws.Range("L2:L" & summaryRow - 1).NumberFormat = "#,##0" ' Total volume format
    Next ws
    
    ' Output the greatest values on the first worksheet
    firstSheet.Cells(2, 14).Value = "Greatest % Increase"
    firstSheet.Cells(3, 14).Value = "Greatest % Decrease"
    firstSheet.Cells(4, 14).Value = "Greatest Total Volume"
    
    ' Add headers for Ticker and Value in columns O1 and P1
    firstSheet.Cells(1, 15).Value = "Ticker"
    firstSheet.Cells(1, 16).Value = "Value"
    
    ' Output the greatest tickers and their values
    firstSheet.Cells(2, 15).Value = greatestIncreaseTicker
    firstSheet.Cells(3, 15).Value = greatestDecreaseTicker
    firstSheet.Cells(4, 15).Value = greatestVolumeTicker
    
    ' Format and display the greatest values
    firstSheet.Cells(2, 16).Value = Format(greatestIncrease * 100, "0.00") & "%" ' Greatest % increase
    firstSheet.Cells(3, 16).Value = Format(greatestDecrease * 100, "0.00") & "%" ' Greatest % decrease
    firstSheet.Cells(4, 16).NumberFormat = "#,##0" ' Greatest total volume formatting
    firstSheet.Cells(4, 16).Value = greatestVolume
    
    ' AutoFit the columns for the greatest values on the first worksheet
    firstSheet.Columns("N:P").AutoFit

End Sub

