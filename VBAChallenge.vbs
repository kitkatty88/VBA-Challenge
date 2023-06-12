Attribute VB_Name = "Module1"
Sub VBAchallenge():
    'Set our Dimensions
    Dim total As Double
    Dim i As Long
    Dim j As Integer
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim start_value As Long
    Dim yearly_change As Double
    Dim percent_change As Double
    
    'Loop through each sheet
        For Each ws In Worksheets
        
            'Define initial variable values
            total = 0
            j = 0
            start_value = 2
            yearly_change = 0
            
        'Set our column titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'Get the last row that has data
        lastrow = Cells(2, 1).End(xlDown).Row
        
            'Start the loop
            For i = 2 To lastrow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    total = total + ws.Cells(i, 7).Value
                
                        'Non zero value
                        If ws.Cells(start_value, 3) = 0 Then
                            For end_value = start_value To i
                                If ws.Cells(end_value, 3).Value <> 0 Then
                                    start_value = end_value
                                    Exit For
                                End If
                            Next end_value
                        Else
                            'zero value
                            If total = 0 Then
                                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                                ws.Range("J" & 2 + j).Value = 0
                                ws.Range("K" & 2 + j).Value = 0
                                ws.Range("L" & 2 + j).Value = 0
                            End If
                        'calculations
                        yearly_change = (ws.Cells(i, 6) - ws.Cells(start_value, 3))
                        percent_change = Round((yearly_change / ws.Cells(start_value, 3) * 100), 2)
                        
                        'keep repeating i
                        start_value = i + 1
                        
                        'print values
                        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                        ws.Range("J" & 2 + j).Value = Round(yearly_change, 2)
                        ws.Range("K" & 2 + j).Value = percent_change & "%"
                        ws.Range("L" & 2 + j).Value = total
                        
                        'Colors
                        If ws.Range("J" & 2 + j).Value < 0 Then
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                            
                            ElseIf ws.Range("J" & 2 + j).Value > 0 Then
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                            
                            Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                            
                        End If
                        
                        End If
                        
                        'reset the initial variable values
                        total = 0
                        j = j + 1
                        yearly_change = 0
                        
                    Else
                        total = total + ws.Cells(i, 7).Value
                    
                End If
            Next i
            
            'Define initial variable values
            increase = 0
            decrease = 0
            volume = 0
            increase_ticker = 0
            decrease_ticker = 0
            volume_ticker = 0
            
            
        'Set our column titles
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
            'Find values for each variable
            For l = 2 To lastrow
                If ws.Range("K" & l).Value > increase Then
                    increase = ws.Range("K" & l).Value
                    increase_ticker = ws.Cells(l, 9).Value
                End If
                If ws.Range("K" & l).Value < decrease Then
                    decrease = ws.Range("K" & l).Value
                    decrease_ticker = ws.Cells(l, 9).Value
                End If
                If ws.Range("L" & l).Value > volume Then
                    volume = ws.Range("L" & l).Value
                    volume_ticker = ws.Cells(l, 9).Value
                End If
            Next l
                
                'print values
                ws.Range("P2").Value = increase * 100 & "%"
                ws.Range("P3").Value = decrease * 100 & "%"
                ws.Range("P4").Value = volume
                ws.Range("O2").Value = increase_ticker
                ws.Range("O3").Value = decrease_ticker
                ws.Range("O4").Value = volume_ticker
                
            
        Next ws


End Sub
