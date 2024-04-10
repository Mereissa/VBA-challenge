Sub year_stock()
    'define demensions (sheets)
    Dim ws As Worksheet
    
    'define loop demension
    Dim i As Long, j As Integer
    
    'define variable deminsion
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalstock As Double
    Dim ticketstart As Double
    
    'add variables for greatest increase, decrease and volume
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatesIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
            
    'reset values of above
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    'define row counting
    Dim rowcount As Long
    
    'start looping throught all sheets
    For Each ws In ThisWorkbook.Worksheets
        With ws
            'output row count
            j = 2
            
            'set title row
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            
            'last row of data
            rowcount = .Cells(.Rows.Count, "A").End(xlUp).Row
            totalstock = 0
            If .Cells(2, 3).Value <> 0 Then
                ticketstart = .Cells(2, 3).Value
            End If
            
            'loop through all rows
            For i = 2 To rowcount
            'check the ticker column is still the same name
                If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
                    'if ticker is diferent do this
                    totalstock = totalstock + .Cells(i, 7).Value
                    yearlyChange = .Cells(i, 6).Value - ticketstart
                    
                    'if ticketstart is not = to zero do this
                    If ticketstart <> 0 Then
                        percentageChange = yearlyChange / ticketstart
                        
                        'if its zero
                    Else: percentageChange = 0
                    End If
                    
                    'check for the greatest increase and decrease
                    If percentageChange > greatestIncrease Then
                        greatestIncrease = percentageChange
                        tickerGreatestIncrease = .Cells(i, 1).Value
                    End If
                    
                    If percentageChange < greatestDecrease Then
                        greatestDecrease = percentageChange
                        tickerGreatestDecrease = .Cells(i, 1).Value
                    End If
                    
                    If totalstock > greatestVolume Then
                        greatestVolume = totalstock
                        tickerGreatestVolume = .Cells(i, 1).Value
                    End If
                    
                    
                    'output the data
                    .Cells(j, 9).Value = .Cells(i, 1).Value
                    .Cells(j, 10).Value = yearlyChange
                    .Cells(j, 10).Interior.Color = IIf(yearlyChange > 0, RGB(0, 255, 0), RGB(255, 0, 0))
                    .Cells(j, 11).Value = percentageChange
                    .Cells(j, 11).NumberFormat = "0.00%"
                    .Cells(j, 12).Value = totalstock
                    
                    'reset the total stock and increment
                    totalstock = 0
                    j = j + 1
                    
                    If .Cells(i + 1, 3).Value <> 0 Then
                    ticketstart = .Cells(i + 1, 3).Value
                    End If
                Else
                   totalstock = totalstock + .Cells(i, 7).Value
                End If
            Next i
            
            'output greatest increase, decrease and volume
            .Cells(2, 14).Value = "Greatest % Increase"
            .Cells(3, 14).Value = "Greatest % Decrease"
            .Cells(4, 14).Value = "Greatest Total Volume"
            .Cells(2, 15).Value = tickerGreatestIncrease
            .Cells(3, 15).Value = tickerGreatestDecrease
            .Cells(4, 15).Value = tickerGreatestVolume
            .Cells(2, 16).Value = greatestIncrease
            .Cells(3, 16).Value = greatestDecrease
            .Cells(4, 16).Value = greatestVolume
            
            'the outcome comes out as percetage
            .Cells(2, 16).NumberFormat = "0.00%"
            .Cells(3, 16).NumberFormat = "0.00%"
            
            
        End With
