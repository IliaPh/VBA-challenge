Attribute VB_Name = "Module1"
Sub MultipleYearStockDataAnalysis():

                                ' # Instructions for Homeworks2
' 1. Create a script that loops through all the stocks for one year and outputs the following information:
' 2. The ticker symbol.
' 3. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' 4. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' 5. The total stock volume of the stock.
' 6. Use conditional formatting that will highlight positive change in green and negative change in red.
    
    
    ' Loop through all sheets
    For Each ws In Worksheets

        ' Create column labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' Set variables
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
        
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        PreviousAmount = 2
        
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

        ' Set the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For I = 2 To LastRow

            ' Add ticker total volume
            TotalTickerVolume = TotalTickerVolume + ws.Cells(I, 7).Value
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

        ' Set ticker names
        TickerName = ws.Cells(I, 1).Value
        ' Print ticker name and total amount in summary table
        ws.Range("I" & SummaryTableRow).Value = TickerName
        ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
        ' Reset ticker total
        TotalTickerVolume = 0


                ' Set Yearly Open, Yearly Close and Yearly Change Name
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & I)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                ' Determine percent change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                ' Format double to include % and two decimal
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                ' Conditional formatting in which positive is green and negative is red
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                ' Add one to summary table
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = I + 1
                End If
            Next I

' Bonus Section: add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    
    '  Adjusting VBA script to allow it to run every worksheet by one click: Added "ws" in a fort of all Range and Cells which will allow to run all worksheets by one click.
        
        'Determine the last row
        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
       
        ' Loop for the final results
        For I = 2 To LastRow
            If ws.Range("K" & I).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & I).Value
                ws.Range("P2").Value = ws.Range("I" & I).Value
            End If

            If ws.Range("K" & I).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & I).Value
                ws.Range("P3").Value = ws.Range("I" & I).Value
            End If

            If ws.Range("L" & I).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & I).Value
                ws.Range("P4").Value = ws.Range("I" & I).Value
            End If

        Next I

    Next ws

End Sub
