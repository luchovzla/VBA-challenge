Sub Stocks():

    ' Define variables
    Dim previous_ticker As String
    Dim current_ticker As String
    Dim next_ticker As String
    Dim ws As Worksheet
    Dim i As Double
    Dim j As Double
    Dim last_row As Double
    Dim open_value As Double
    Dim close_value As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim volume As Double
    Dim total_volume As Double
    Dim color_positive As Integer
    Dim color_negative As Integer
    Dim current_percentage As Double
    Dim current_volume As Double
    Dim max_percentage As Double
    Dim min_percentage As Double
    Dim max_volume As Double
    Dim max_ticker As String
    Dim min_ticker As String
    Dim max_volume_ticker As String

' For loop to cycle between worksheets

For Each ws In Worksheets
    
        ' Last Row
        last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Create headers and define counter for loops
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        i = 2
        j = 2
        
        ' Define color indexes to be used
        color_positive = 10 'Darker green so it doesn't hurt the eyes
        color_negative = 30 'Darker red
    
        For i = 2 To last_row
    
            ' Read tickers
            previous_ticker = ws.Cells(i - 1, 1).Value
            current_ticker = ws.Cells(i, 1).Value
            next_ticker = ws.Cells(i + 1, 1).Value
            
            ' Cumulative addition of volumes
            volume = ws.Cells(i, 7).Value
            total_volume = total_volume + volume
    
            ' If statement to store year open value
            If current_ticker <> previous_ticker Then
                open_value = ws.Cells(i, 3).Value
            End If
    
            ' If statement to close ticker and store variables
            If next_ticker <> current_ticker Then
            
                ' Write current ticker on column I
                ws.Cells(j, 9).Value = current_ticker
                
                ' Store year close value in variable
                close_value = ws.Cells(i, 6).Value
                                
                ' Define yearly change and write in column J
                yearly_change = open_value - close_value
                ws.Cells(j, 10).Value = yearly_change
                ws.Cells(j, 10).NumberFormat = "#.00"
                
                ' If statement to format color fill of Yearly Change cell
                If yearly_change > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = color_positive
                ElseIf yearly_change < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = color_negative
                End If
                
                ' Write Percent Change in column K
                
                If open_value <> 0 And close_value <> 0 Then
                    percent_change = (open_value - close_value) / open_value
                    ws.Cells(j, 11).Value = percent_change
                    ws.Cells(j, 11).NumberFormat = "#.00%"
                Else
                    percent_change = 0
                    ws.Cells(j, 11).Value = percent_change
                    ws.Cells(j, 11).NumberFormat = "#.00%"
                End If
                
                ' Write total stock volume in column L
                ws.Cells(j, 12).Value = total_volume
                
                ' Reset counters
                j = j + 1
                total_volume = 0
            End If
    
        Next i
    
    ' Now let's go through the summary to find the highest % increase, lowest % decrease
    ' and highest total stocks moved
    
    last_row = ws.Cells(Rows.Count, "I").End(xlUp).Row
    ' Init max and min percentages and highest volume
    max_percentage = ws.Cells(2, 11).Value
    min_percentage = ws.Cells(2, 11).Value
    max_volume = ws.Cells(2, 12).Value
    
    ' Summary table headers
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ' For loop to look up max and min values
    
    For i = 2 To last_row
    
        ' Read values from current cell and next cell
        current_percentage = ws.Cells(i, 11).Value
        current_volume = ws.Cells(i, 12).Value
        
        ' Compare and store values if condition met for highest % increase
        If current_percentage > max_percentage Then
            max_percentage = current_percentage
            max_ticker = ws.Cells(i, 9).Value
        End If
        
        ' Compare and store values if condition met for lowest % decrease
        If current_percentage < min_percentage Then
            min_percentage = current_percentage
            min_ticker = ws.Cells(i, 9).Value
        End If
        
        ' Compare and store values if condition met for highest stock volume
        If current_volume > max_volume Then
            max_volume = current_volume
            max_volume_ticker = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    ' Populate summary table
    ws.Range("P2").Value = max_ticker
    ws.Range("P3").Value = min_ticker
    ws.Range("P4").Value = max_volume_ticker
    ws.Range("Q2").Value = max_percentage
    ws.Range("Q2").NumberFormat = "#.00%"
    ws.Range("Q3").Value = min_percentage
    ws.Range("Q3").NumberFormat = "#.00%"
    ws.Range("Q4").Value = max_volume
    
    Next ws

End Sub