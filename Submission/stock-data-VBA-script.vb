
Sub stockTrackerAllSheets()

    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4") ' list of sheet names of process
    
    Dim lastRow As Long
    Dim i As Long ' Current row number
    Dim k As Long ' Row in columns I and J to write results
    Dim stockName As String
    Dim startRow As Long ' First row of a stock's data
    Dim endRow As Long ' Last row of a stock's data
    Dim openingPrice As Double ' Opening price from the first row of a stock's data
    Dim closingPrice As Double ' Closing price from the last row of a stock's data
    Dim quarterlyChange As Double ' Difference between closing and opening prices
    Dim percentChange As Double ' quarterlyChange/openingPrice * 100
    Dim totalVolume As Double ' total volume for stock
    
    ' Variables for tracking maximum and minimum percentage changes and volumes
    Dim maxIncrease As Double
    Dim maxIncreaseName As String
    Dim minDecrease As Double
    Dim minDecreaseName As String
    Dim maxVolume As Double
    Dim maxVolumeName As String

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)  ' Specify the correct sheet name
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row  ' Finds the last row with data in column A

        maxIncrease = -1E+308  ' Initialize to a very small number
        minDecrease = 1E+308   ' Initialize to a very large number
        maxVolume = 0          ' Initialize to zero
        startRow = 2  ' first row of data to process
        k = 2  ' Start writing results from row 2 in tracker table
    
        For i = 2 To lastRow
            ' Check if the current row's stock name is different from the next row's or it's the last row
            If i = lastRow Or ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                stockName = ws.Cells(startRow, 1).Value
                endRow = i
                openingPrice = ws.Cells(startRow, 3).Value
                closingPrice = ws.Cells(endRow, 6).Value
                quarterlyChange = closingPrice - openingPrice
                percentChange = quarterlyChange / openingPrice
                
                'Sum volume for this stock
                totalVolume = 0
                Dim volumeRow As Long
                For volumeRow = startRow To endRow
                    totalVolume = totalVolume + ws.Cells(volumeRow, 7).Value ' sum column G for each stock
                Next volumeRow
                
                ' Check for max increase, min decrease, and max volume
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseName = stockName
                End If
                If percentChange < minDecrease Then
                    minDecrease = percentChange
                    minDecreaseName = stockName
                End If
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeName = stockName
                End If
    
                ' Write results to the change tracker table
                ws.Cells(k, 9).Value = stockName  ' write stock name in column I
                ws.Cells(k, 10).Value = quarterlyChange  ' write quarterly change in column J
                
                    'conditional formatting
                    If (quarterlyChange > 0) Then
                        ws.Cells(k, 10).Interior.ColorIndex = 4 ' green
                    ElseIf (quarterlyChange < 0) Then
                        ws.Cells(k, 10).Interior.ColorIndex = 3 ' red
                    Else
                        ws.Cells(k, 10).Interior.ColorIndex = 2 ' white
                    End If

                ws.Cells(k, 11).Value = percentChange ' write percent change in column K
                ws.Cells(k, 11).NumberFormat = "0.00%" ' shows as % with 2 decimal places
                ws.Cells(k, 12).Value = totalVolume ' write total stock volue in column L
                
                ' Go to next row in tracker table to write the next result in the next row of columns I and J
                k = k + 1
    
                ' Update startRow for the next stock block
                startRow = i + 1
                
            End If
        Next i

        ' Write max increase, min decrease, and max volume to columns O and P
        ws.Cells(2, 15).Value = maxIncreaseName
        ws.Cells(2, 16).Value = maxIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%" ' shows as % with 2 decimal places
        ws.Cells(3, 15).Value = minDecreaseName
        ws.Cells(3, 16).Value = minDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%" ' shows as % with 2 decimal places
        ws.Cells(4, 15).Value = maxVolumeName
        ws.Cells(4, 16).Value = maxVolume
    Next sheetName
    
End Sub


