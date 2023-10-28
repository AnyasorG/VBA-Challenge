Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

    Dim Year As Integer
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim ResultRow As Long
    Dim Ticker As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim TickerIncrease As String
    Dim TickerDecrease As String
    Dim TickerVolume As String

    ' To Loop through the years (worksheets)
    For Year = 2018 To 2020 ' Update the range of years as needed
        Set ws = ThisWorkbook.Worksheets(CStr(Year))
 
            ' Calculate the last row
            LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row


            ' Initialize variables for each sheet
            ResultRow = 2 ' Data start from row 2
            Ticker = ws.Cells(2, 1).Value
            YearOpen = ws.Cells(2, 3).Value
            YearClose = 0
            YearlyChange = 0
            PercentChange = 0
            TotalVolume = 0

            TickerIncrease = ""
            TickerDecrease = ""
            TickerVolume = ""

            
            For i = 2 To LastRow
                
                If Ticker = ws.Cells(i, 1).Value Then
                    
                    YearClose = ws.Cells(i, 6).Value ' Closing price for the current date
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value ' Accumulate volume
                Else
                    ' Calculate Yearly Change and Percentage Change for the previous stock
                    YearlyChange = YearClose - YearOpen
                    If YearOpen <> 0 Then
                        PercentChange = ((YearClose - YearOpen) / YearOpen) * 100
                    Else
                        PercentChange = 0 '
                    End If

                    ' Check for greatest % increase
                    If PercentChange > MaxIncrease Then
                        MaxIncrease = PercentChange
                        TickerIncrease = Ticker
                    End If

                    ' Check for greatest % decrease
                    If PercentChange < MaxDecrease Then
                        MaxDecrease = PercentChange
                        TickerDecrease = Ticker
                    End If

                    ' Check for greatest total volume
                    If TotalVolume > MaxVolume Then
                        MaxVolume = TotalVolume
                        TickerVolume = Ticker
                    End If

                    ws.Cells(1, 9).Value = "Ticker"
                    ws.Cells(1, 10).Value = "Yearly Change"
                    ws.Cells(1, 11).Value = "Percent Change"
                    ws.Cells(1, 12).Value = "Total Stock Volume"
                    ws.Cells(ResultRow, 9).Value = Ticker ' Column 9 is for Ticker
                    ws.Cells(ResultRow, 10).Value = YearlyChange ' Column 10 is for Yearly Change
                    ws.Cells(ResultRow, 11).Value = Format(PercentChange / 100, "0.00%") ' Format Percentage Change

                    ' Format the color of cells in columns J (YearlyChange) and K (PercentChange)
                    If YearlyChange > 0 Then
                        ws.Cells(ResultRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive YearlyChange
                    ElseIf YearlyChange < 0 Then
                        ws.Cells(ResultRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative YearlyChange
                    Else
                        ws.Cells(ResultRow, 10).Interior.ColorIndex = xlNone ' No color for zero YearlyChange
                    End If

                    If PercentChange > 0 Then
                        ws.Cells(ResultRow, 11).Interior.Color = RGB(0, 255, 0) ' Green for positive PercentChange
                    ElseIf PercentChange < 0 Then
                        ws.Cells(ResultRow, 11).Interior.Color = RGB(255, 0, 0) ' Red for negative PercentChange
                    Else
                        ws.Cells(ResultRow, 11).Interior.ColorIndex = xlNone ' No color for zero PercentChange
                    End If

                    ws.Cells(ResultRow, 12).Value = TotalVolume ' Column 12 is for Total Stock Volume

                    Ticker = ws.Cells(i, 1).Value
                    YearOpen = ws.Cells(i, 3).Value
                    YearClose = ws.Cells(i, 6).Value
                    YearlyChange = 0
                    PercentChange = 0
                    TotalVolume = ws.Cells(i, 7).Value
                    ResultRow = ResultRow + 1
                End If
            Next i

            
            YearlyChange = YearClose - YearOpen 'Yearly Change for the last stock
            If YearOpen <> 0 Then
                PercentChange = ((YearClose - YearOpen) / YearOpen) * 100
            Else
                PercentChange = 0 ' Percentage Change for the last stock
            End If

            ' Check for greatest % increase, greatest % decrease, and greatest total volume
            If PercentChange > MaxIncrease Then
                MaxIncrease = PercentChange
                TickerIncrease = Ticker
            End If

            If PercentChange < MaxDecrease Then
                MaxDecrease = PercentChange
                TickerDecrease = Ticker
            End If

            If TotalVolume > MaxVolume Then
                MaxVolume = TotalVolume
                TickerVolume = Ticker
            End If

         
            ws.Cells(ResultRow, 9).Value = Ticker ' Column 9 is for Ticker
            ws.Cells(ResultRow, 10).Value = YearlyChange ' Column 10 is for Yearly Change
            ws.Cells(ResultRow, 11).Value = Format(PercentChange / 100, "0.00%") ' Format Percentage Change

            ' Format the color of cells in columns J (Yearly Change) and K (Percent Change)
            If YearlyChange > 0 Then
                ws.Cells(ResultRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive YearlyChange
            ElseIf YearlyChange < 0 Then
                ws.Cells(ResultRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative YearlyChange
            Else
                ws.Cells(ResultRow, 10).Interior.ColorIndex = xlNone ' No color for zero YearlyChange
            End If

            If PercentChange > 0 Then
                ws.Cells(ResultRow, 11).Interior.Color = RGB(0, 255, 0) ' Green for positive PercentChange
            ElseIf PercentChange < 0 Then
                ws.Cells(ResultRow, 11).Interior.Color = RGB(255, 0, 0) ' Red for negative PercentChange
            Else
                ws.Cells(ResultRow, 11).Interior.ColorIndex = xlNone ' No color for zero PercentChange
            End If

            ws.Cells(ResultRow, 12).Value = TotalVolume ' Column 12 is for Total Stock Volume

            ' Move the Greatest values to Column O and P
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(2, 16).Value = TickerIncrease
            ws.Cells(3, 16).Value = TickerDecrease
            ws.Cells(4, 16).Value = TickerVolume

            ' Format the percentages in Column Q (Rows 2 and 3)
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 17).Value = Format(MaxIncrease / 100, "0.00%")
            ws.Cells(3, 17).Value = Format(MaxDecrease / 100, "0.00%")
            ws.Cells(4, 17).Value = MaxVolume
        
    Next Year
End Sub




