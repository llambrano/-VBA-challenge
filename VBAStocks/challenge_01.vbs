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

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("J:J").NumberFormat = "0.00"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("K:K").NumberFormat = "0.00%"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"

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

        ElseIf Cells(c, 1).Value <> Cells(c + 1, 1).Value Then
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

            Case Is >= 0
                Range("J" & locationSumTab).Interior.Color = 11854022
            
            Case Is < 0
                Range("J" & locationSumTab).Interior.Color = 10202109

            End Select

            ' move to the next row
            locationSumTab = locationSumTab + 1

        Else
            ' capture volume
            volume = volume + Cells(c, 7).Value

        End If

    Next c

    ' mac and min percentages and volume 
    Range("P2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("P3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("P4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    perMax = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    perMin = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    vol = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' adding ticker to summary table
    Range("O2") = Cells(perMax + 1, 9)
    Range("O3") = Cells(perMin + 1, 9)
    Range("O4") = Cells(vol + 1, 9)

End Sub