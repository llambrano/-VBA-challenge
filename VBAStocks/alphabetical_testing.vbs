Sub StockMarket():

' Set an initial variable for holding the ticker name
    Dim TickerName As String

' Set an initial variable for holding the total per ticker

    Dim VolumeTotal As Double
    volume_total = 0

' Keep track of the location for each ticker in the summary table

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

' Identify last row

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through ticker quotes

    For i = 2 To lastrow

' Loop to calculate total volume and summarize tickers
' Check if we are still within the same ticker, if it is not...

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

        ' Set the ticker name
        TickerName = Cells(i, 1).Value

        ' Print the ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = TickerName

        ' Add to the ticker Total
        VolumeTotal = VolumeTotal + Cells(i, 7).Value
        
        ' Print the volume total to the Summary Table
        Range("L" & Summary_Table_Row).Value = VolumeTotal

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset the Brand Total
        VolumeTotal = 0
        
        ClosingPrice = Cells(i, 6)
        Range("N" & Summary_Table_Row - 1).Value = ClosingPrice

        ' If the cell immediately following a row is the same ticker...
        
        Elseif Cells(i, 1).Value <> Cells(i - 1, 1).Value Then

        OpeningPrice = Cells(i, 3)
        Range("M" & Summary_Table_Row).Value = OpeningPrice

        
        ' Add to the ticker Total
        VolumeTotal = Cells(i, 7).Value

        Else 
        VolumeTotal = VolumeTotal + Cells(i, 7)
           
    End If

    Next i

End Sub




        OpeningPrice = Cells(i, 3)

        Range("J" & Summary_Table_Row).Value = ClosingPrice - OpeningPrice
        Range("K" & Summary_Table_Row).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
        Range("M" & Summary_Table_Row).Value = OpeningPrice
        Range("N", & Summary_Table_Row).Value = ClosingPrice - 1

Dim close_price As Double
    Dim open_price As Double
    Dim yearlychange As Integer

Y = 2

    For x = 1 To 80000
        If Cells(x + 1, 1).Value = Cells(x + 2, 1) Then
            If Cells(x + 1, 2).Value = 20160101 Then
                open_price = Cells(x + 1, 3).Value
            End If
        Else
             close_price = Cells(x + 1, 6).Value
             Cells(Y, 9).Value = Cells(x + 1, 1).Value
             Cells(Y, 10).Value = close_price - open_price
             Cells(Y, 11).Value = (close_price - open_price) / open_price
             Y = 1 + Y
        End If
    Next x

End Sub
