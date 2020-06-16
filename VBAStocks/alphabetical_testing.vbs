Sub tickers()

    ' declare vars
    Dim ticker As String
    Dim c As Long
    Dim locationSumTab As Integer
    Dim volume As Double
    Dim openingPrice As Single
    Dim closingPrice As Single
    Dim percentageChange As Single
    Dim netChange As Single
    Dim j As Integer

    ' location summary tab 
    locationSumTab = 2
    volume = 0

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
            'Range("N" & locationSumTab).Value = openingPrice

        Elseif Cells(c, 1).Value <> Cells(c + 1, 1).Value Then
            closingPrice = Cells(c, 6).Value
            ' Range("O" & locationSumTab).Value = closingPrice

            ' calculate change 
            netChange = closingPrice - openingPrice
            Range("J" & locationSumTab).Value = netChange

            ' calculate percentage change
            percentageChange = (closingPrice - openingPrice) / openingPrice
            Range("K" & locationSumTab).Value = percentageChange

            ' capture volume
            volume = volume + Cells(c, 7).Value
            Range("L" & locationSumTab).Value = volume
            
            ' reset volume
            volume = 0

            ' color netChamnge
            Select Case netChange

            Case is => 0
                Range("J" & locationSumTab).Interior.Color = 10202109
            
            Case is < 0
                Range("J" & locationSumTab).Interior.Color = 11854022

            End Select

            ' move to the next row
            locationSumTab = locationSumTab + 1

        Else
            ' capture volume
            volume = volume + Cells(c, 7).Value

        End If

    Next c

        ' take the max and min and place them in a separate part in the worksheet
    Range("P2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("P3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("P4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    ' returns one less because header row not a factor
    volMax = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volMin = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    vol = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' final ticker symbol for  total, greatest % of increase and decrease, and average
    Range("O2") = Cells(volMax + 1, 9)
    Range("O3") = Cells(volMin + 1, 9)
    Range("O4") = Cells(vol + 1, 9)

End Sub


