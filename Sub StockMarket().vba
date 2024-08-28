Sub StockMarket()
    ' Loop through all worksheets
    For Each ws In Worksheets
        ws.Activate

        ' Add Variables
        Dim lastRow As Long
        Dim Ticker As String
        Dim SummaryTableRow As Integer
        Dim i As Long
        Dim Openprice As Double
        Dim Closeprice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim Volume As Double
        Dim SummaryTable As Long
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestTotalVolume As Double

        ' Add Counters
        Volume = 0
        SummaryTable = 2

        ' Find lastRow in worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Create Column headers
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Quarterly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"

        ' Add Greatest Increase, Greatest Decrease, and Greatest total volume headers
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"

        ' Capture lastRow Loop
        For i = 2 To lastRow
            ' Ticker
            If ws.Cells(i, "A").Value <> ws.Cells(i - 1, "A").Value Then
                ' Set Ticker Name
                Ticker = ws.Cells(SummaryTable, "I").Value
                ws.Cells(SummaryTable, "I").Value = ws.Cells(i, "A").Value
                
                ' Set Open Price
                Openprice = ws.Cells(i, "C").Value
            
            ElseIf ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
                ' Set Yearly Change
                YearlyChange = ws.Cells(i, "F").Value - Openprice
                ws.Cells(SummaryTable, "J").Value = YearlyChange

                ' Set Percent Change
                If Openprice <> 0 Then
                    PercentChange = YearlyChange / Openprice
                Else
                    PercentChange = 0
                End If
                
                ws.Cells(SummaryTable, "K").Value = PercentChange
                ws.Cells(SummaryTable, "K").NumberFormat = "0.00%"

                ' Apply Conditional Formatting
                If YearlyChange >= 0 Then
                    ws.Cells(SummaryTable, "J").Interior.Color = vbGreen
                Else
                    ws.Cells(SummaryTable, "J").Interior.Color = vbRed
                End If

                If PercentChange >= 0 Then
                    ws.Cells(SummaryTable, "K").Interior.Color = vbGreen
                Else
                    ws.Cells(SummaryTable, "K").Interior.Color = vbRed
                End If

                ' Set Volume
                ws.Cells(SummaryTable, "L").Value = Volume
                Volume = 0
                SummaryTable = SummaryTable + 1
            Else
                Volume = Volume + ws.Cells(i, "G").Value
            End If
        Next i

        ' Add functionality to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        GreatestIncrease = Application.WorksheetFunction.Max(ws.Columns("K"))
        GreatestDecrease = Application.WorksheetFunction.Min(ws.Columns("K"))
        GreatestTotalVolume = Application.WorksheetFunction.Max(ws.Columns("L"))

        ' Set "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q4").Value = GreatestTotalVolume
        
        ' Format "Greatest % increase" and "Greatest % decrease" as percentages
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

        ' Use XLOOKUP to find the corresponding tickers
        ws.Range("P2").Value = Application.WorksheetFunction.XLookup(GreatestIncrease, ws.Columns("K"), ws.Columns("I"))
        ws.Range("P3").Value = Application.WorksheetFunction.XLookup(GreatestDecrease, ws.Columns("K"), ws.Columns("I"))
        ws.Range("P4").Value = Application.WorksheetFunction.XLookup(GreatestTotalVolume, ws.Columns("L"), ws.Columns("I"))
        
        ' Auto fit columns for better visibility
        ws.Columns("I:Q").AutoFit
        
    Next ws
End Sub
