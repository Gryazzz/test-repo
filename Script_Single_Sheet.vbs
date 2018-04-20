'Any level can be chosen: Easy, Moderate of Hard

Sub Easy()

    Cells(1, 9).Value = "Ticker"
    Cells(1, 12).Value = "Total Stock Volume"

    dim CurrTicker as Variant, NextTicker as Variant
    dim k as LongLong, RowCount as LongLong
    dim Volume as LongLong

    RowCount = Cells(Rows.Count, 1).End(xlUp).Row
    k = 2
    Volume = 0

    for i = 2 to RowCount
        CurrTicker = Cells(i,1).value
        NextTicker = Cells(i+1,1).value
        if CurrTicker = NextTicker Then
            Volume = Volume + Cells(i,7).value
        else
            Volume = Volume + Cells(i,7).value
            Cells(k,9).value = CurrTicker
            Cells(k,12).value = Volume
            Volume = 0
            k = k + 1
        end if
    next i

End Sub


sub Moderate()

    Cells(1,10).value = "Yearly Change"
    Cells(1,11).value = "Percent Change"
    
    Call Easy

    Dim RowCount As LongLong
    dim Start as Double, Closing as Double
    dim Before as LongLong, After as LongLong
    dim k as LongLong

    RowCount = Cells(Rows.Count, 1).End(xlUp).Row
    Start = 0
    Closing = 0
    k = 2
    Range("J:J").NumberFormat = "0.00"
    Range("K:K").NumberFormat = "0.00%"

    for i = 2 to RowCount
        Before = i - 1
        After = i + 1

        if Cells(Before,1).value <> Cells(i,1).value Then
           Start = Start + Cells(i,3).value
        end if

        if Cells(After,1).value <> Cells(i,1).value Then
            Closing = Closing + Cells(i,6).value
            Cells(k,10).value = Closing - Start
            
            if Start <> 0 Then
                Cells(k,11).value = Closing / Start - 1
            else
                Start = 1
                Cells(k,11).value = Closing / Start - 1
            end if 

            if Cells(k,10).value >= 0 Then 'Coloring the cells
                Cells(k,10).Interior.ColorIndex = 4
            else
                Cells(k,10).Interior.ColorIndex = 3
            end if

            k = k + 1
            Closing = 0
            Start = 0
        end if

    next i
    
end sub


sub Hard()

    Cells(2,15).value = "Greatest % Increase"
    Cells(3,15).value = "Greatest % Decrease"
    Cells(4,15).value = "Greatest Total Volume"
    Cells(1,16).value = "Ticker"
    Cells(1,17).value = "Value"
    
    Call Moderate

    Dim RowCount As LongLong
    dim rg1 as Range, rg2 as Range
    dim minValue as Double, maxValue as Double, maxTot as Double

    Set rg1 = Range("K:K")
    Set rg2 = Range("L:L")
    RowCount = Cells(Rows.Count, 11).End(xlUp).Row
    minValue = Application.WorksheetFunction.Min(rg1)
    maxValue = Application.WorksheetFunction.Max(rg1)
    maxTot = Application.WorksheetFunction.Max(rg2)
    
    Cells(2,17).NumberFormat = "0.00%"
    Cells(3,17).NumberFormat = "0.00%"
    Cells(2,17).value = maxValue
    Cells(3,17).value = minValue
    Cells(4,17).value = maxTot

    For i = 2 to RowCount 'search the tickers
        if Cells(i,11).value = maxValue Then
            Cells(2,16).value = Cells(i,9).value
        end if
        if Cells(i,11).value = minValue Then
            Cells(3,16).value = Cells(i,9).value
        end if
        if Cells(i,12).value = maxTot Then
            Cells(4,16).value = Cells(i,9).value
        end if
    next i

end sub

sub Unique() 'Callback, List of unique tickers, Don't need it anymore
    Cells(2,9).value = Cells(2,1).value
    dim k as LongLong
    k = 2
    for i = 2 to Cells(Rows.Count, 1).End(xlUp).Row
        if Cells(i,1).value <> Cells(k,9).value Then
            k = k + 1
            Cells(k,9).value = Cells(i,1)
        end if
    next i    
end sub
