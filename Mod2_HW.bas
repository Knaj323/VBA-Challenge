Attribute VB_Name = "Module1"
Sub hw()
Dim lastRow As Long
Dim ticker As String
Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentageChange As Double
Dim totalVolume As Double
Dim summaryrow As Long
Dim ws As Worksheet
Dim greatestIncreaseTicker As String
Dim greatestIncreasePercentage As Double
Dim greatestDecreaseTicker As String
Dim greatestDecreasePercentage As Double
Dim greatestTotalVolumeTicker As String
Dim greatestTotalVolume As Double


Set wb = ThisWorkbook
For Each ws In wb.Sheets
    
summaryrow = 2
greatestIncreasePercentage = 0
greatestDecreasePercentage = 0
greatestTotalVolume = 0
    
  
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Change"

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        closingPrice = ws.Cells(i, 6).Value
        
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            openingPrice = ws.Cells(i, 3).Value
            yearlyChange = closingPrice - openingPrice
            If openingPrice <> 0 Then
                percentageChange = (closingPrice / openingPrice)
            Else
                percentageChange = 0
            End If
            totalVolume = ws.Cells(i, 7).Value
            
            If percentageChange > greatestIncreasePercentage Then
            greatestIncreasePercentage = percentageChange
            greatestIncreaseTicker = ticker
        End If

        If percentageChange < greatestDecreasePercentage Then
            greatestDecreasePercentage = percentageChange
            greatestDecreaseTicker = ticker
        End If

        If totalVolume > greatestTotalVolume Then
            greatestTotalVolume = totalVolume
            greatestTotalVolumeTicker = ticker
        End If
            
            
            ws.Cells(summaryrow, 9).Value = ticker
            ws.Cells(summaryrow, 10).Value = yearlyChange
            ws.Cells(summaryrow, 11).Value = percentageChange
            ws.Cells(summaryrow, 11).NumberFormat = "0.00%"
            ws.Cells(summaryrow, 11).Value = percentageChange
            ws.Cells(summaryrow, 12).Value = totalVolume
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(2, 16).Value = greatestIncreaseTicker
            ws.Cells(3, 16).Value = greatestDecreaseTicker
            ws.Cells(4, 16).Value = greatestTotalVolumeTicker
            ws.Cells(2, 17).Value = greatestIncreasePercentage
            ws.Cells(3, 17).Value = greatestDecreasePercentage
            ws.Cells(4, 17).Value = greatestTotalVolume
            
            If yearlyChange >= 0 Then
                ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
            End If
            
           
            summaryrow = summaryrow + 1
        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    Next i
Next ws
End Sub
