Sub StockMarket():

Dim Y As Integer
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
