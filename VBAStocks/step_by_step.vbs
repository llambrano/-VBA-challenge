Sub tickers()

    ' declare vars
    Dim ticker As String
    Dim c As Long
    Dim locationSumTab As Integer
    Dim volume As Double
    Dim openingPrice As Integer
    Dim closingPrice As Integer
    Dim percentageChange As Single
    Dim netChange As Single

    ' location summary tab 
    locationSumTab = 2

    ' identify the range 
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    ' get tickers | Identify if previous ticker is <>
    For c = 2 To rowCount
        If Cells(c, 1).Value <> Cells(c - 1, 1).Value Then
        
            ' add ticker to summary table 
            ticker = Cells(c, 1).Value
            Range("I" & locationSumTab).Value = ticker

            ' capture opening price to summary table
            openingPrice = Cells(c, 3).Value
            Range("N" & locationSumTab).Value = openingPrice


        ElseIf Cells(c, 1).Value <> Cells(c + 1, 1).Value Then
            closingPrice = Cells(c, 6).Value
            Range("O" & locationSumTab).Value = closingPrice

            ' calculate change 
            netChange = closingPrice - openingPrice
            Range("J" & locationSumTab).Value = netChange

            ' calculate percentage change
            percentageChange = (closingPrice - openingPrice) / openingPrice
            Range("K" & locationSumTab).Value = percentageChange

            ' move to the next row
            locationSumTab = locationSumTab + 1
        End If
    Next c

End Sub





