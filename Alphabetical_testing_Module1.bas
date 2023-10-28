Attribute VB_Name = "Module1"

Sub Alphabetical_testing()
    
    
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

    ' Loop through the years (worksheets)
    For Each ws In ThisWorkbook.Worksheets
        
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
       

          
            For i = 2 To LastRow
                
                If Ticker = ws.Cells(i, 1).Value Then
                    YearClose = ws.Cells(i, 6).Value ' Closing price for the current date
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                Else
                    ' Calculate Yearly Change and Percentage Change for the previous stock
                    YearlyChange = YearClose - YearOpen
                    If YearOpen <> 0 Then
                        PercentChange = ((YearClose - YearOpen) / YearOpen) * 100
                    Else
                        PercentChange = 0
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

                    
                    ws.Cells(ResultRow, 9).Value = Ticker ' Column 9 is for Ticker
                    ws.Cells(ResultRow, 10).Value = YearlyChange ' Column 10 is for Yearly Change
                    ws.Cells(ResultRow, 11).Value = Format(PercentChange / 100, "0.00%") ' Format Percentage Change
                    ws.Cells(ResultRow, 12).Value = TotalVolume ' Column 12 is for Total Stock Volume

                    ' Format the color to Yearly Change (Column J)
                    If YearlyChange > 0 Then
                        ws.Cells(ResultRow, 10).Interior.Color = RGB(0, 255, 0) ' Green fill for positive values
                    ElseIf YearlyChange < 0 Then
                        ws.Cells(ResultRow, 10).Interior.Color = RGB(255, 0, 0) ' Red fill for negative values
                    End If

                    ' Format Color to Percent Change (Column K)
                    If PercentChange > 0 Then
                        ws.Cells(ResultRow, 11).Interior.Color = RGB(0, 255, 0) ' Green fill for positive values
                    ElseIf PercentChange < 0 Then
                        ws.Cells(ResultRow, 11).Interior.Color = RGB(255, 0, 0) ' Red fill for negative values
                    End If

                    Ticker = ws.Cells(i, 1).Value
                    YearOpen = ws.Cells(i, 3).Value
                    YearClose = ws.Cells(i, 6).Value
                    YearlyChange = 0
                    PercentChange = 0
                    TotalVolume = ws.Cells(i, 7).Value
                    ResultRow = ResultRow + 1
                End If
            Next i

            ' Calculate Yearly Change and Percentage Change
            YearlyChange = YearClose - YearOpen
            If YearOpen <> 0 Then
                PercentChange = ((YearClose - YearOpen) / YearOpen) * 100
            Else
                PercentChange = 0
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

            
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"

            ' Move the Greatest values to Column O
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(2, 16).Value = TickerIncrease
            ws.Cells(3, 16).Value = TickerDecrease
            ws.Cells(4, 16).Value = TickerVolume

            ' Format the percentages in Column Q (Rows 2 and 3)
            ws.Cells(2, 17).Value = Format(MaxIncrease / 100, "0.00%")
            ws.Cells(3, 17).Value = Format(MaxDecrease / 100, "0.00%")
            ws.Cells(4, 17).Value = MaxVolume
        
    Next ws
End Sub







