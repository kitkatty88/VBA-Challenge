# VBA-Challenge

Created with the help of peers and TA outside of office hours

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
                            
Created with the help of TA outside of office hours
increase_ticker = ws.Cells(l, 9).Value
decrease_ticker = ws.Cells(l, 9).Value
volume_ticker = ws.Cells(l, 9).Value
