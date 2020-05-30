'-----------------------------------------------------------------------
' Name: Tameka Kuar
' Date: 5/30/2020
' In addition to my own code, I used credit_card_challenge.vbs,
' wells_fargo_ptl.vbs, and lotto_numbers.vbs as a guide for
' this subroutine VBAStocks()
'-----------------------------------------------------------------------

Sub VBAStocks():

    '--------------------------------------------------------------
    ' For Each Loop to Loop through every Worksheet
    '--------------------------------------------------------------
    For Each ws In Worksheets

        Dim TickerSymbol As String
        Dim BeginningPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim SummaryRow As Integer
        Dim LastRow As Long

        BeginningPrice = ws.Cells(2, 3).Value
        SummaryRow = 2
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        '-----------------------------------
        '            HEADERS
        '-----------------------------------
        ws.Range("I1 , P1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        '--------------------------------------
        '    Loop to fill Summary Table
        '--------------------------------------
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Sets the ticker symbol
                TickerSymbol = ws.Cells(i, 1).Value
                
                ' Calculates the yearly change
                YearlyChange = ws.Cells(i, 6).Value - BeginningPrice
            
                ' Avoids dividing by 0
                If BeginningPrice = 0 Then
                    
                    PercentChange = 0
                
                Else
                    
                    ' Calculates the percent change
                    PercentChange = YearlyChange / BeginningPrice
                
                End If
            
                ' Adds to the total stock volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

                ' Prints the ticker symbol in the summary table
                ws.Range("I" & SummaryRow).Value = TickerSymbol
                
                ' Prints the yearly change in the summary table
                ws.Range("J" & SummaryRow).Value = YearlyChange
                
                ' Prints the percent change in the summary table
                ws.Range("K" & SummaryRow).Value = PercentChange
                
                ' Prints the total stock volume to the summary table
                ws.Range("L" & SummaryRow).Value = TotalVolume
                
                ' Adds one to the summary table row
                SummaryRow = SummaryRow + 1
                
                ' Sets the beginning opening price to the new ticker symbol's
                BeginningPrice = ws.Cells(i + 1, 3).Value
                
                ' Reset the total stock volume
                TotalVolume = 0
                
            ' If the cell immediately following the previous row
            ' has the same ticker symbol then the following will occur
            Else
        
            ' Adds to the total stock volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
            End If
            
        Next i
        
        '-----------------------------------
        '            FORMATTING
        '-----------------------------------
        Dim SummaryLast As Long

        ' The last row in the summary table
        SummaryLast = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        For i = 2 To SummaryLast
            ws.Cells(i, 10).NumberFormat = "0.00"
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
            
        Next i
        
        For i = 2 To SummaryLast
            ws.Cells(i, 11).NumberFormat = "0.00%"
        Next i
        '----------------------------------------
        ' Loop for the Second Summary Table
        '----------------------------------------

        Dim GreatInc As Double
        Dim GreatDec As Double
        Dim GreatVol As Double
        Dim GreatIncTic As String
        Dim GreatDecTic As String
        Dim GreatVolTic As String

        GreatInc = ws.Cells(2, 11).Value
        GreatDec = ws.Cells(2, 11).Value
        GreatVol = ws.Cells(2, 12).Value

        For i = 2 To SummaryLast

            If ws.Cells(i, 11).Value > GreatInc Then
               
                GreatInc = ws.Cells(i, 11).Value
                GreatIncTic = ws.Cells(i, 9).Value

            End If
        

            If (ws.Cells(i, 11).Value < GreatDec) Then
                
                GreatDec = ws.Cells(i, 11).Value
                GreatDecTic = ws.Cells(i, 9).Value
            
            End If

            If ws.Cells(i, 12).Value > GreatVol Then
                
                GreatVol = ws.Cells(i, 12).Value
                GreatVolTic = ws.Cells(i, 9).Value
            
            End If

        Next i
        
        ' Prints values in the Second Summary Table
        ws.Cells(2, 16).Value = GreatIncTic
        ws.Cells(3, 16).Value = GreatDecTic
        ws.Cells(4, 16).Value = GreatVolTic

        ws.Cells(2, 17).Value = GreatInc
        ws.Cells(3, 17).Value = GreatDec
        ws.Cells(4, 17).Value = GreatVol
        '-----------------------------------
        '         MORE FORMATTING
        '-----------------------------------
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Columns("A:Q").AutoFit
        
    Next ws
    
End Sub
