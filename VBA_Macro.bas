Attribute VB_Name = "Module1"
Sub Analysis():

    ' Loop Through All Worksheets
    For Each ws In Worksheets

        'Identify Column Headers for (Ticker, Yearly change, Percent change, Toatl stock Volume, Greatest Increase, Greatest Decrease, Ticker and Value respectlly.
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        'Declare all needed Variables
        Dim Ticker_name As String
        Dim Yearly_Change, Summary_table, previuos_amount As Long
        Dim Percent_Change, Yearly_Open, Yearly_Close, Total_Ticker_Volume As Double
        Dim Greatest_Increase, Greatest_Decrease, Greatest_Total_Volume As Double
        Dim last_raw As Double
        Total_Ticker_Volume = 0
        Summary_table_row = 2
        Previous_amount = 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        Greatest_Total_Volume = 0
        

        ' look for the Last Row in worksheets
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' some vlues won't fit, so we do auto fit
        ws.Columns("I:Q").AutoFit
        'loop for total ticker
        For i = 2 To last_row

            ' Add To Ticker Total Volume
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            'we should be in the same tiker_name, otherwise..
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' try to Set Ticker_name
                Ticker_name = ws.Cells(i, 1).Value
                ' now Print our Ticker_name In Summary_table
                ws.Range("I" & Summary_table_row).Value = Ticker_name
                ' Print Ticker_total_amount Summary_table
                ws.Range("L" & Summary_table_row).Value = Total_Ticker_Volume
                ' Reset Ticker Total
                Total_Ticker_Volume = 0

                ' find Yearly_Open, Yearly_Close and Yearly_Change Name
                Yearly_Open = ws.Range("C" & Previous_amount)
                Yearly_Close = ws.Range("F" & i)
                Yearly_Change = Yearly_Close - Yearly_Open
                ws.Range("J" & Summary_table_row).Value = Yearly_Change

                ' find our Percent Change
                If Yearly_Open = 0 Then
                    Percent_Change = 0
                Else
                    Yearly_Open = ws.Range("C" & Previous_amount)
                    Percent_Change = Yearly_Change / Yearly_Open
                End If
                ' Formatting to % Symbol And creating Two Decimal Places, tm make it Double
                ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
                ws.Range("K" & Summary_table_row).Value = Percent_Change

                '  Highlight the Positive change to Green color and Negative change to Red color
                If ws.Range("J" & Summary_table_row).Value >= 0 Then
                    ws.Range("J" & Summary_table_row).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Range("J" & Summary_table_row).Interior.Color = RGB(255, 0, 0)
                End If
                'to make sure we go through this loop, add one to sammury table and previuos amount
            
                Summary_table_row = Summary_table_row + 1
                Previous_amount = i + 1
                End If
            Next i
            'count Greatest increase an decrease, alongside greateset total

            last_row = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' loop for study results
            For i = 2 To last_row
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        'formatting to get % symbol and 2 Dicimal places
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
        

    Next ws

End Sub
