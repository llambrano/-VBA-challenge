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

    ' define range to find highest and lowest
    Set percRange = Worksheets("A").Range("K:K")
    
    ' finding max % change
    perMax = Excel.WorksheetFunction.Max(percRange)
    Range("N2") = "Greatest Percentage Increse"
    Range("O2").NumberFormat = ("0.0%")
    Range("O2") = perMax
    
    ' finding min % change
    perMin = Excel.WorksheetFunction.Min(percRange)
    Range("N3") = "Greatest Percentage Decrese"
    Range("O3").NumberFormat = ("0.0%")
    Range("O3") = perMin

    ' finding greatest total volume
    Set volRange = Worksheets("A").Range("L:L")

    maxVol = Excel.WorksheetFunction.Max(volRange)
    Range("N4") = "Greatest Total Volume"
    Range("O4").NumberFormat = ("#,##0_ ;-#,##0")
    Range("O4") = maxVol

End Sub


