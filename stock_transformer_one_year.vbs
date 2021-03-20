Function transform_stock_data()
    Dim year
    
    ' Set Year Here
    year = "2014"

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

    LastRow =Worksheets(year).Cells(Rows.Count, 1).End(xlUp).Row
    Worksheets(year).Cells(1,9).Value = "Ticker"
    Worksheets(year).Cells(1,10).Value = "Yearly Change"
    Worksheets(year).Cells(1,11).Value = "Percent Change"
    Worksheets(year).Cells(1,12).Value = "Total Stock Volume"
    currRow = 2
    For i=2 to LastRow
        ' Handles starter case where currStock is empty
        if Len(currStock) = 0 Then
            currStock =Worksheets(year).Cells(i,1)
            currStockStartPrice =Worksheets(year).Cells(i,3)
            currStockEndPrice =Worksheets(year).Cells(i,6)
            currStockTotalVol =Worksheets(year).Cells(i,7)
        ' Standard case where we're on the same stock as before
        ElseIf StrComp(currStock, Worksheets(year).Cells(i,1)) = 0 Then
            currStockEndPrice =Worksheets(year).Cells(i,6)
            currStockTotalVol = currStockTotalVol +Worksheets(year).Cells(i,7)
        Else
            Worksheets(year).Cells(currRow,9).Value = currStock
            Worksheets(year).Cells(currRow,10).Value = currStockEndPrice - currStockStartPrice 
            if currStockEndPrice - currStockStartPrice > 0 Then
                Worksheets(year).Cells(currRow, 10).Interior.ColorIndex = 4
            Else 
                Worksheets(year).Cells(currRow, 10).Interior.ColorIndex = 3
            End if
            if currStockStartPrice =0 Then
                Worksheets(year).Cells(currRow,11).Value = "Undefined"
            Else
                Worksheets(year).Cells(currRow,11).Value = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
                if (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice) > greatestIncrease Then
                    greatestIncrease = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
                    greatestIncreaseStock = currStock
                End If
                if (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice) < greatestDecrease Then
                    greatestDecrease = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
                    greatestDecreaseStock = currStock
                End If
            End if
            Worksheets(year).Cells(currRow,11).NumberFormat = "0.00%"
            Worksheets(year).Cells(currRow,12).Value = currStockTotalVol
            if currStockTotalVol - greatestVolume > 0 Then
                greatestVolume = currStockTotalVol
                greatestVolumeStock = currStock
            end if
            currRow = currRow + 1
            currStock =Worksheets(year).Cells(i,1)
            currStockStartPrice =Worksheets(year).Cells(i,3)
            currStockEndPrice =Worksheets(year).Cells(i,6)
            currStockTotalVol =Worksheets(year).Cells(i,7)
        End if
    Next
    Worksheets(year).Cells(currRow,9).Value = currStock
    Worksheets(year).Cells(currRow,10).Value = currStockEndPrice - currStockStartPrice 
    if currStockEndPrice - currStockStartPrice > 0 Then
        Worksheets(year).Cells(currRow, 10).Interior.ColorIndex = 4
    Else 
        Worksheets(year).Cells(currRow, 10).Interior.ColorIndex = 3
    End if
    if currStockStartPrice =0 Then
        Worksheets(year).Cells(currRow,11).Value = "Undefined"
    Else
        Worksheets(year).Cells(currRow,11).Value = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
        if (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice) > greatestIncrease Then
            greatestIncrease = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
            greatestIncreaseStock = currStock
        End If
        if (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice) < greatestDecrease Then
            greatestDecrease = (currStockEndPrice - currStockStartPrice)/ Abs(currStockStartPrice)
            greatestDecreaseStock = currStock
        End If
    End if
    Worksheets(year).Cells(currRow,11).NumberFormat = "0.00%"
    Worksheets(year).Cells(currRow,12).Value = currStockTotalVol
    if currStockTotalVol - greatestVolume > 0 Then
        greatestVolume = currStockTotalVol
        greatestVolumeStock = currStock
    end if
    currStock = ""
    Worksheets(year).Cells(1,16).Value = "Ticker"
    Worksheets(year).Cells(1,17).Value = "Value"
    Worksheets(year).Cells(2,15).Value = "Greatest % Increase"
    Worksheets(year).Cells(2,16).Value = greatestIncreaseStock
    Worksheets(year).Cells(2,17).Value = greatestIncrease
    Worksheets(year).Cells(2,17).NumberFormat = "0.00%"
    Worksheets(year).Cells(3,15).Value = "Greatest & Decrease"
    Worksheets(year).Cells(3,16).Value = greatestDecreaseStock
    Worksheets(year).Cells(3,17).Value = greatestDecrease
    Worksheets(year).Cells(3,17).NumberFormat = "0.00%"
    Worksheets(year).Cells(4,15).Value = "Greatest Total Volume"
    Worksheets(year).Cells(4,16).Value = greatestVolumeStock
    Worksheets(year).Cells(4,17).Value = greatestVolume
End Function
    
