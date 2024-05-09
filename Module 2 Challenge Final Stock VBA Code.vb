Sub stock_loop()

    For Each ws In Worksheets

            ' Name the ticker column that is going to be the variable (stock letters)
            Dim ticker_name As String

            ' Set an initial variable for holding the total ticker stock volume
            Dim total_stock_volume As Double
            total_stock_volume = 0

            ' Keep track of the location for each ticker in the summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2

            ' Name the other variables going to be used in the "if" statement
            Dim opening_price As Double
            Dim closing_price As Double
            Dim quarterly_change As Double
            Dim percent_change As Double

            ' Identify the new headers will be located (Ticker, Quarterly Change, Percent Change, Total Stock Volume)
            ws.Range("I1").Value = "Ticker"
                ws.Range("I1").Interior.ColorIndex = 6
                ws.Range("I1").Font.Bold = True
    
            ws.Range("J1").Value = "Quarterly Change"
                ws.Range("J1").Interior.ColorIndex = 6
                ws.Range("J1").Font.Bold = True

            ws.Range("K1").Value = "Percent Change"
                ws.Range("K1").Interior.ColorIndex = 6
                ws.Range("K1").Font.Bold = True
    
            ws.Range("L1").Value = "Total Stock Volume"
                ws.Range("L1").Interior.ColorIndex = 6
                ws.Range("L1").Font.Bold = True

            ' There are way too many rows, so we will define what the last row is (determine last row)
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' Loop through all ticker names
            For i = 2 To LastRow

                ' Check if we are still within the same ticker name, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                    ' Set the ticker name
                    ticker_name = ws.Cells(i, 1).Value
        
                    ' Add the total stock volume
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
                    ' Print the ticker name in the Summary Table
                    ws.Range("I" & Summary_Table_Row).Value = ticker_name
        
                    ' Print the total stock volume in the Summary Table
                    ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
        
                    ' Identify the initial closing and opening prices variables
                    closing_price = ws.Cells(i, 6).Value
                    opening_price = ws.Cells(i, 3).Value                    
                    
                    'Calculate the quarterly change
                    quarterly_change = closing_price - opening_price

                    ' Print the value of the quarterly change for each ticker name
                    ws.Range("J" & Summary_Table_Row).Value = quarterly_change

                    'Calculate the percent change
                    percent_change = quarterly_change / opening_price

                    ' Print the percent change for each ticker
                    ws.Range("K" & Summary_Table_Row).Value = percent_change

                        ' Format the percent change to a percent
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"                    
                    
                    ' Add one to the Summary Table Row (to allow the ticker to change)
                    Summary_Table_Row = Summary_Table_Row + 1
        
                    '   Reset the total stock volume (to reset the values for the next ticker)
                    total_stock_volume = 0
        
                Else
        
                    ' Add the volume of trade stock volume if the ticker name is the same
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
                End If

            Next i


            ' Make a loop to color the quarterly change 

            LastRow_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row

            For i = 2 To LastRow_Summary_Table
    
                If ws.Cells(i, 10).Value > 0 Then
        
                    ws.Cells(i, 10).Interior.ColorIndex = 4 'Green
    
                ElseIf ws.Cells(i, 10).Value < 0 Then
    
                    ws.Cells(i, 10).Interior.ColorIndex = 3 'Red
        
                End If
        
            Next i

            For i = 2 To LastRow_Summary_Table
        
                If ws.Cells(i, 11).Value > 0 Then
    
                    ws.Cells(i, 11).Interior.ColorIndex = 4 'Green
            
                ElseIf ws.Cells(i, 11).Value < 0 Then
        
                    ws.Cells(i, 11).Interior.ColorIndex = 3 'Red
        
                End If
    
            Next i

            ' Label the appropriate cells like in the image
            ws.Range("P1").Value = "Ticker"
                ws.Range("P1").Interior.ColorIndex = 6
                ws.Range("P1").Font.Bold = True
    
            ws.Range("Q1").Value = "Value"
                ws.Range("Q1").Interior.ColorIndex = 6
                ws.Range("Q1").Font.Bold = True
    
            ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O2").Interior.ColorIndex = 6
                ws.Range("O2").Font.Bold = True
    
            ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O3").Interior.ColorIndex = 6
                ws.Range("O3").Font.Bold = True
    
            ws.Range("O4").Value = "Greatest Total Volume"
                ws.Range("O4").Interior.ColorIndex = 6
                ws.Range("O4").Font.Bold = True


            ' For loop to go through all the date
            For i = 2 To LastRow_Summary_Table

                If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow_Summary_Table)) Then
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                    ws.Range("Q2").Value = ws.Cells(i, 11).Value
                        ws.Range("Q2").NumberFormat = "0.00%"
        
                ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow_Summary_Table)) Then
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                    ws.Range("Q3").Value = ws.Cells(i, 11).Value
                        ws.Range("Q3").NumberFormat = "0.00%"
    
                ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow_Summary_Table)) Then
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                    ws.Range("Q4").Value = ws.Cells(i, 12).Value
    
                End If
    
            Next i

    Next ws

End Sub






















