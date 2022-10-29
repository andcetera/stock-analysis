Sub StockTicker()

    'create basic variables
    Dim TotVol As Double
    Dim Opn As Double
    Dim Change As Double
    Dim RowNum As Integer
   
    'vars for greatest percent increase/decrease and total volume
    Dim Gpi as Double
    Dim GpiTick as String
    Dim Gpd as Double
    Dim GpdTick as String
    Dim Gtv as Double
    Dim GtvTick as String
   
    For Each ws in Worksheets
        'borrwed variable from census_pt1 project solved to get last row of worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'set var starting levels for each worksheet
        RowNum = 2
        Gpi = 0
        Gpd = 0
        Gtv = 0

        'print table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        'loop thru all rows & check 1st column for ticker symbol
        For i = 2 To LastRow

            'add every column 7 for running total volume
            TotVol = TotVol + ws.Cells(i, 7).Value

            'Simple If: get row 1 column 3 for opening (first row only) - store in var
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                Opn = ws.Cells(i, 3).Value
            End If

            'when next column 1 is different, get that row's column 6 for close and print out info in summary chart
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'print ticker symbol (column A) in column 9
                ws.Cells(RowNum, 9).Value = ws.Cells(i, 1).Value

                'change = close - open, print in column 10 & format color r/g
                Change = ws.Cells(i, 6).Value - Opn
                ws.Cells(RowNum, 10).Value = Change
                If Change >= 0 Then
                    ws.Cells(RowNum, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(RowNum, 10).Interior.ColorIndex = 3
                End If

                'print percentage change & format in column 11
                ws.Cells(RowNum, 11).Value = FormatPercent(Change/Opn)

                'check for greatest increase/decrease percentages
                If ws.Cells(RowNum, 11).Value > Gpi Then
                    Gpi = ws.Cells(RowNum, 11).Value
                    GpiTick = ws.Cells(i, 1).Value
                ElseIf ws.Cells(RowNum, 11).Value < Gpd Then
                    Gpd = ws.Cells(RowNum, 11).Value
                    GpdTick = ws.Cells(i, 1).Value
                End If

                'print running total volume in column 12
                ws.Cells(RowNum, 12).Value = TotVol

                'increase row number for summary chart
                RowNum = RowNum + 1

                'check greatest total volume before zeroing out
                If TotVol > Gtv Then
                    Gtv = TotVol
                    GtvTick = ws.Cells(i, 1).Value
                End If

                'reset running total var
                TotVol = 0
            End If

        Next i

        'print greatest increase/decrease & total volume for whole sheet
        ws.Cells(2, 15).Value = GpiTick
        ws.Cells(2, 16).Value = FormatPercent(Gpi)
        ws.Cells(3, 15).Value = GpdTick
        ws.Cells(3, 16).Value = FormatPercent(Gpd)
        ws.Cells(4, 15).Value = GtvTick
        ws.Cells(4, 16).Value = Gtv

    Next ws

End Sub