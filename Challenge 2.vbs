Attribute VB_Name = "Module1"
ub stock_data():
    'Create Variables
    Dim WorksheetName As String
    Dim year_change As Double
    Dim percent As Double
    Dim sum_table As Integer
    Dim maxIncreaseCode As String
    Dim maxDecreaseCode As String
    Dim maxVolumeCode As String
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As LongLong
    Dim total As LongLong
    Dim start As Long
    
    For Each ws In Worksheets
        'Determine the last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Grabbed the WorksheetName
        WorksheetName = ws.Name
    
    year_change = 0
    start = 2
    sum_table = 2
    total = 0
    
    'Set variables to track greatest values
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    
    'Create columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
        'Loop through the rows to calculate yearly change, percent change and total volume of each stock
        For i = 2 To lastRow
    
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                         
                                            
               'Calculate yearly change & percent change
               year_change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
               percent = year_change / ws.Cells(start, 3).Value
               
               
               'Calculate total volume
               total = total + ws.Cells(i, 7).Value
               
               'Output the information in summary table
               ws.Range("I" & sum_table).Value = ws.Cells(i, 1).Value
               ws.Range("J" & sum_table).Value = year_change
               ws.Range("K" & sum_table).Value = percent
               ws.Range("K" & sum_table).NumberFormat = "0.00%"
               ws.Range("L" & sum_table).Value = total
               
               'Autofit the columns
               ws.Range("I" & sum_table).EntireColumn.AutoFit
               ws.Range("J" & sum_table).EntireColumn.AutoFit
               ws.Range("K" & sum_table).EntireColumn.AutoFit
               ws.Range("L" & sum_table).EntireColumn.AutoFit
               ws.Range("O1:Q4").EntireColumn.AutoFit
               
                    'Apply conditional formatting
                    If year_change >= 0 And percent >= 0 Then
                        ws.Range("J" & sum_table).Interior.ColorIndex = 4
                        ws.Range("K" & sum_table).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & sum_table).Interior.ColorIndex = 3
                        ws.Range("K" & sum_table).Interior.ColorIndex = 3
                    End If
                    
                    'Find the stock with the greatest percentage increase
                    If percent > maxIncrease Then
                       maxIncrease = percent
                       maxIncreaseCode = ws.Cells(i, 1).Value
                    End If
                    
                    'Find the stock with the greatest percentage decrease
                    If percent < maxDecrease Then
                       maxDecrease = percent
                       maxDecreaseCode = ws.Cells(i, 1).Value
                    End If
                    
                    'Find the stock with the greatest total volume
                    If total > maxVolume Then
                       maxVolume = total
                       maxVolumeCode = ws.Cells(i, 1).Value
                    
                    'Output greatest values into columns
                    ws.Cells(2, 15).Value = "Greatest % Increase"
                    ws.Cells(3, 15).Value = "Greatest % Decrease"
                    ws.Cells(4, 15).Value = "Greatest Total Volume"
                    ws.Range("P1").Value = "Ticker"
                    ws.Range("Q1").Value = "Value"
                    ws.Range("P2").Value = maxIncreaseCode
                    ws.Range("P3").Value = maxDecreaseCode
                    ws.Range("P4").Value = maxVolumeCode
                    ws.Range("Q2").Value = maxIncrease
                    ws.Range("Q3").Value = maxDecrease
                    ws.Range("Q4").Value = maxVolume
                    End If
                   
                'Reset the values
                sum_table = sum_table + 1
                total = 0
        
            Else
                total = total + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
    Next ws

End Sub




