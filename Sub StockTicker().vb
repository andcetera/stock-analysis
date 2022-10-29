Sub StockTicker()
    'create basic variables
    Dim TotVol As Double
    Dim Opn As Double
    Dim Change As Double
    Dim RowNum As Integer
    RowNum = 2

    'vars for greatest percent increase/decrease and total volume
    Dim Gpi as Double
    Gpi = 0
    Dim GpiTick as String
    Dim Gpd as Double
    Gpd = 0
    Dim GpdTick as String
    Dim Gtv as Double
    Gtv = 0
    Dim GtvTick as String

    'var for each worksheet
    Dim ws as String
    'borrwed variable from census_pt1 project solved to get last row of worksheet
    LastRow = Sheet1.Cells(Rows.Count, 1).End(xlUp).Row

    'print table headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"

    'loop thru all rows & check 1st column for ticker symbol
    For i = 2 To LastRow

        'add every column 7 for running total volume
        TotVol = TotVol + Cells(i, 7).Value

        'Simple If: get row 1 column 3 for opening (first row only) - store in var
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            Opn = Cells(i, 3).Value
        End If

        'when next column 1 is different, get that row's column 6 for close and print out info in summary chart
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'print ticker symbol (column A) in column 9
            Cells(RowNum, 9).Value = Cells(i, 1).Value

            'change = close - open, print in column 10 & format color r/g
            Change = Cells(i, 6).Value - Opn
            Cells(RowNum, 10).Value = Change
            If Change >= 0 Then
                Cells(RowNum, 10).Interior.ColorIndex = 4
            Else
                Cells(RowNum, 10).Interior.ColorIndex = 3
            End If

            'print percentage change & format in column 11
            Cells(RowNum, 11).Value = FormatPercent(Change/Opn)

            'check for greatest increase/decrease percentages
            If Cells(RowNum, 11).Value > Gpi Then
                Gpi = Cells(RowNum, 11).Value
                GpiTick = Cells(i, 1).Value
            ElseIf Cells(RowNum, 11).Value < Gpd Then
                Gpd = Cells(RowNum, 11).Value
                GpdTick = Cells(i, 1).Value
            End If

            'print running total volume in column 12
            Cells(RowNum, 12).Value = TotVol

            'increase row number for summary chart
            RowNum = RowNum + 1

            'check greatest total volume before zeroing out
            If TotVol > Gtv Then
                Gtv = TotVol
                GtvTick = Cells(i, 1).Value
            End If

            'reset running total var
            TotVol = 0
        End If
    Next i

    'print greatest increase/decrease & total volume for whole sheet
    Cells(2, 15).Value = GpiTick
    Cells(2, 16).Value = FormatPercent(Gpi)
    Cells(3, 15).Value = GpdTick
    Cells(3, 16).Value = FormatPercent(Gpd)
    Cells(4, 15).Value = GtvTick
    Cells(4, 16).Value = Gtv

End Sub