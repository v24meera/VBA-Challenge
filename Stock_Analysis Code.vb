Sub Stock_Analysis()
    Dim ws As Worksheet
    Dim Last_Row As Long
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim Summary_Row As Long
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
    Dim Ticker_Increase As String
    Dim Ticker_Decrease As String
    Dim Ticker_Volume As String
    
    ' For loop to loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ' Calling the last row in column A
        Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Summary table variables
        Summary_Row = 2
        Total_Volume = 0
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Volume = 0
        
        ' Headers for summary table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        ' For loop to loop through the data
        For i = 2 To Last_Row
            ' Check if ticker is different from the previous one
            If Cells(i, 1).Value <> Ticker Then
                If Total_Volume <> 0 Then
                    ' Yearly and percent change calculations
                    Yearly_Change = Close_Price - Open_Price
                    If Open_Price <> 0 Then
                        Percent_Change = (Yearly_Change / Open_Price) * 100
                    Else
                        Percent_Change = 0
                    End If
                    
                    ' Summary table output
                    Cells(Summary_Row, 9).Value = Ticker
                    Cells(Summary_Row, 10).Value = Yearly_Change
                    Cells(Summary_Row, 11).Value = Percent_Change
                    Cells(Summary_Row, 12).Value = Total_Volume
                    
                    ' Conditional formatting to Yearly Change column
                    If Yearly_Change > 0 Then
                        Cells(Summary_Row, 10).Interior.Color = RGB(0, 255, 0)
                    ElseIf Yearly_Change < 0 Then
                        Cells(Summary_Row, 10).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    '  Setting and updating greatest increase/decrease/volume metrics
                    If Percent_Change > Greatest_Increase Then
                        Greatest_Increase = Percent_Change
                        Ticker_Increase = Ticker
                    End If
                    If Percent_Change < Greatest_Decrease Then
                        Greatest_Decrease = Percent_Change
                        Ticker_Decrease = Ticker
                    End If
                    If Total_Volume > Greatest_Volume Then
                        Greatest_Volume = Total_Volume
                        Ticker_Volume = Ticker
                    End If
                    
                    ' Resetting variables for the new ticker
                    Summary_Row = Summary_Row + 1
                    Total_Volume = 0
                End If
                
                ' Setting new ticker and open price
                Ticker = Cells(i, 1).Value
                Open_Price = Cells(i, 3).Value
            End If
            
            ' Summing total volume
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
            ' Setting close price for the current ticker
            Close_Price = Cells(i, 6).Value
        Next i
        
        ' Getting yearly and percent change for the last ticker
        Yearly_Change = Close_Price - Open_Price
        If Open_Price <> 0 Then
            Percent_Change = (Yearly_Change / Open_Price) * 100
        Else
            Percent_Change = 0
        End If
        
        ' Output for the last ticker to the summary table
        Cells(Summary_Row, 9).Value = Ticker
        Cells(Summary_Row, 10).Value = Yearly_Change
        Cells(Summary_Row, 11).Value = Percent_Change
        Cells(Summary_Row, 12).Value = Total_Volume
        
        ' Conditional formatting for yearly Change column for the last ticker
        If Yearly_Change > 0 Then
            Cells(Summary_Row, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf Yearly_Change < 0 Then
            Cells(Summary_Row, 10).Interior.Color = RGB(255, 0, 0)
        End If
        
        ' Updating greatest % increase, decrease and total volume for the last ticker
        If Percent_Change > Greatest_Increase Then
            Greatest_Increase = Percent_Change
            Ticker_Increase = Ticker
        End If
        If Percent_Change < Greatest_Decrease Then
            Greatest_Decrease = Percent_Change
            Ticker_Decrease = Ticker
        End If
        If Total_Volume > Greatest_Volume Then
            Greatest_Volume = Total_Volume
            Ticker_Volume = Ticker
        End If
        
        ' Summary table outputs for greatest values
        Cells(2, 16).Value = Ticker_Increase
        Cells(2, 17).Value = Greatest_Increase & "%"
        Cells(3, 16).Value = Ticker_Decrease
        Cells(3, 17).Value = Greatest_Decrease & "%"
        Cells(4, 16).Value = Ticker_Volume
        Cells(4, 17).Value = Greatest_Volume
    Next ws
End Sub
