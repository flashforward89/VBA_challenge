' Homework#2
'Create a script that will loop through all the stocks for one year and output the following information.
Sub SMAnalysis():

    ' Setting loop for all WS
    For Each ws In Worksheets

        ' Titles of columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        'Declare variables
        'Ticker Name
        Dim TName As String
        'Last Row
        Dim LRow As Double
        'Total Ticker Volume
        Dim TTVolume As Double
        TTVolume = 0
        'Summary Table Row
        Dim STRow As Double
        STRow = 2
        'Year Close
        Dim YClose As Double
        'Year Open
        Dim YOpen As Double
       'Year Change
        Dim YChange As Double
        Dim PreviousAmount As Double
        PreviousAmount = 2
        Dim PercentChange As Double
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim LastRowValue As Double
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

        ' Last Row looping // you could use range or cell option to define specific cells
        LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LRow
                     
            ' Suming up Ticker Total Volume
            TTVolume = TTVolume + ws.Cells(i, 7).Value
            ' Check the same Ticker else
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Set Ticker Name
            TName = ws.Cells(i, 1).Value
            ws.Range("L" & STRow).Value = TTVolume
            ws.Range("I" & STRow).Value = TName
            
            ' Reset Ticker
             TTVolume = 0

                'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
                ' Set yearly open
                YOpen = ws.Range("C" & PreviousAmount)
                'Set yearly close
                YClose = ws.Range("F" & i)
                'Set yearly change
                YChange = YClose - YOpen
                ws.Range("J" & STRow).Value = YChange

                ' Cutting the loop if year open is zero
                If YOpen = 0 Then
                    PercentChange = 0
                Else
                    YOpen = ws.Range("C" & PreviousAmount)
                    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
                    PercentChange = YChange / YOpen
                End If
                ' Format for % and to decimals
                ws.Range("K" & STRow).NumberFormat = "0.00%"
                ws.Range("K" & STRow).Value = PercentChange

                ' You should also have conditional formatting that will highlight positive change in green and negative change in red.
                If ws.Range("J" & STRow).Value >= 0 Then
                    ws.Range("J" & STRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & STRow).Interior.ColorIndex = 3
                End If
                STRow = STRow + 1
                PreviousAmount = i + 1
                End If
            Next i
            
            'BONUS

            'Loop
            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Greatest % Increase
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If
            'Greatest % Decreased
                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If
            'Greatest total volumn
                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        ' Format decimal places for %
            ws.Range("Q2, Q3").NumberFormat = "0.00%"
           

    Next ws

End Sub