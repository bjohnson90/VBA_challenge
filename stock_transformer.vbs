Function transform_stock_data()
    Dim currStock As String
    Dim greatestIncreaseStock As String
    Dim greatestDecreaseStock As String
    Dim greatestVolumeStock AS String
    Dim currStockStartPrice 
    Dim currStockEndPrice  
    Dim currStockPercentChange 
    Dim currStockTotalVol 
    Dim workSheetCount
    Dim greatestIncrease
    Dim greatestDecrease
    Dim greatestVolume
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

    ' Below stores row for writing results
    Dim currRow
    workSheetCount = Worksheets.Count

    For w=1 To workSheetCount
        LastRow =Worksheets(w).Cells(Rows.Count, 1).End(xlUp).Row
        Worksheets(w).Cells(1,9).Value = "Ticker"
        Worksheets(w).Cells(1,10).Value = "Yearly Change"
        Worksheets(w).Cells(1,11).Value = "Percent Change"
        Worksheets(w).Cells(1,12).Value = "Total Stock Volume"
        currRow = 2
        For i=2 to LastRow
            ' Handles starter case where currStock is empty
            if Len(currStock) = 0 Then
                currStock =Worksheets(w).Cells(i,1)
                currStockStartPrice =Worksheets(w).Cells(i,3)
                currStockEndPrice =Worksheets(w).Cells(i,6)
                currStockTotalVol =Worksheets(w).Cells(i,7)
            ' Standard case where we're on the same stock as before
            ElseIf StrComp(currStock, Worksheets(w).Cells(i,1)) = 0 Then
                currStockEndPrice =Worksheets(w).Cells(i,6)
                currStockTotalVol = currStockTotalVol +Worksheets(w).Cells(i,7)
            Else
                Worksheets(w).Cells(currRow,9).Value = currStock
                Worksheets(w).Cells(currRow,10).Value = currStockEndPrice - currStockStartPrice 
                if currStockEndPrice - currStockStartPrice > 0 Then
                    Worksheets(w).Cells(currRow, 10).Interior.ColorIndex = 4
                Else 
                    Worksheets(w).Cells(currRow, 10).Interior.ColorIndex = 3
                End if
                if currStockStartPrice =0 Then
                    Worksheets(w).Cells(currRow,11).Value = "Undefined"
                Else
                    Worksheets(w).Cells(currRow,11).Value = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
                    if (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice) > greatestIncrease Then
                        greatestIncrease = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
                        greatestIncreaseStock = currStock
                    End If
                    if (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice) < greatestDecrease Then
                        greatestDecrease = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
                        greatestDecreaseStock = currStock
                    End If
                End if
                Worksheets(w).Cells(currRow,11).NumberFormat = "0.00%"
                Worksheets(w).Cells(currRow,12).Value = currStockTotalVol
                if currStockTotalVol - greatestVolume > 0 Then
                    greatestVolume = currStockTotalVol
                    greatestVolumeStock = currStock
                end if
                currRow = currRow + 1
                currStock =Worksheets(w).Cells(i,1)
                currStockStartPrice =Worksheets(w).Cells(i,3)
                currStockEndPrice =Worksheets(w).Cells(i,6)
                currStockTotalVol =Worksheets(w).Cells(i,7)
            End if
        Next
        Worksheets(w).Cells(currRow,9).Value = currStock
        Worksheets(w).Cells(currRow,10).Value = currStockEndPrice - currStockStartPrice 
        if currStockEndPrice - currStockStartPrice > 0 Then
            Worksheets(w).Cells(currRow, 10).Interior.ColorIndex = 4
        Else 
            Worksheets(w).Cells(currRow, 10).Interior.ColorIndex = 3
        End if
        if currStockStartPrice =0 Then
            Worksheets(w).Cells(currRow,11).Value = "Undefined"
        Else
            Worksheets(w).Cells(currRow,11).Value = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
            if (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice) > greatestIncrease Then
                greatestIncrease = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
                greatestIncreaseStock = currStock
            End If
            if (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice) < greatestDecrease Then
                greatestDecrease = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
                greatestDecreaseStock = currStock
            End If
        End if
        Worksheets(w).Cells(currRow,11).NumberFormat = "0.00%"
        Worksheets(w).Cells(currRow,12).Value = currStockTotalVol
        if currStockTotalVol - greatestVolume > 0 Then
            greatestVolume = currStockTotalVol
            greatestVolumeStock = currStock
        end if
        currStock = ""
    Next
    Worksheets(1).Cells(1,16).Value = "Ticker"
    Worksheets(1).Cells(1,17).Value = "Value"
    Worksheets(1).Cells(2,15).Value = "Greatest % Increase"
    Worksheets(1).Cells(2,16).Value = greatestIncreaseStock
    Worksheets(1).Cells(2,17).Value = greatestIncrease
    Worksheets(1).Cells(2,17).NumberFormat = "0.00%"
    Worksheets(1).Cells(3,15).Value = "Greatest & Decrease"
    Worksheets(1).Cells(3,16).Value = greatestDecreaseStock
    Worksheets(1).Cells(3,17).Value = greatestDecrease
    Worksheets(1).Cells(3,17).NumberFormat = "0.00%"
    Worksheets(1).Cells(4,15).Value = "Greatest Total Volume"
    Worksheets(1).Cells(4,16).Value = greatestVolumeStock
    Worksheets(1).Cells(4,17).Value = greatestVolume
End Function
    
