Sub YearlyChange()
    For Each ws In Worksheets 
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlup).Row


        Dim ticker As String 
        Dim BeginPrice As Double 
        Dim EndPrice As Double 
        Dim volume As Double 
        Dim maxinc, maxdec, maxvol As Double 
        Dim maxincticker, maxdecticker, maxvolticker As String
        

        ticker = ws.Cells(2, 1).Value
        BeginPrice = ws.Cells(2, 6).Value
        volume = ws.Cells(2, 7).Value

        Dim recordidx As Integer
        recordidx = 2



        For i = 2 To lastrow
            If ws.Cells(i, 1) <> ticker Then


                ws.Cells(recordidx, 9).Value = ticker
                ws.Cells(recordidx, 10).Value = EndPrice - BeginPrice
                If ws.Cells(recordidx, 10).Value >= 0 Then
                    ws.Cells(recordidx, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(recordidx, 10).Interior.ColorIndex = 3
                End If
                
                If BeginPrice > 0 Then
                    ws.Cells(recordidx, 11).Value = (EndPrice- BeginPrice) / BeginPrice
                Else
                    ws.Cells(recordidx, 11).Value = 0.0
                End If
                
                ws.Cells(recordidx, 12).Value = volume
            
                If recordidx = 2 Then
                    maxinc = ws.Cells(recordidx, 11).Value
                    maxincticker = ticker
                    maxdec = ws.Cells(recordidx, 11).Value
                    maxdecticker = ticker
                    maxvol = volume
                    maxvolticker = ticker
                Else
                    If ws.Cells(recordidx, 11).Value > maxinc Then
                        maxinc = ws.Cells(recordidx, 11).Value
                        maxincticker = ticker
                    ElseIf ws.Cells(recordidx, 11).Value < maxdec Then
                        maxdec = ws.Cells(recordidx, 11).Value
                        maxdecticker = ticker
                    End If
                    
                    If recordidx = 2 Or volume > maxvol Then
                        maxvol = volume
                        maxvolticker = ticker
                    End If
                End If

                recordidx = recordidx + 1
            
                ticker = ws.Cells(i, 1).Value
                BeginPrice = ws.Cells(i, 6).Value
                volume = ws.Cells(i, 7).Value
            Else
                EndPrice = ws.Cells(i, 6).Value
                volume = volume + ws.Cells(i, 7).Value
            End If
        
            If i = lastrow Then
                ws.Cells(recordidx, 9).Value = ticker
                ws.Cells(recordidx, 10).Value = EndPrice - BeginPrice
                If ws.Cells(recordidx, 10).Value >= 0 Then
                    ws.Cells(recordidx, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(recordidx, 10).Interior.ColorIndex = 3
                End If

                If BeginPrice > 0 Then
                    ws.Cells(recordidx, 11).Value = (EndPrice - BeginPrice) / BeginPrice
                Else
                    ws.Cells(recordidx, 11).Value = 0.0
                End If
                
                ws.Cells(recordidx, 12).Value = volume
            
                If ws.Cells(recordidx, 11).Value > maxinc Then
                    maxinc = ws.Cells(recordidx, 11).Value
                    maxincticker = ticker
                ElseIf ws.Cells(recordidx, 11).Value < maxdec Then
                    maxdec = ws.Cells(recordidx, 11).Value
                    maxdecticker = ticker
                End If
                    
                If recordidx = 2 Or volume > maxvol Then
                    maxvol = volume
                    maxvolticker = ticker
                End If
            End If
        Next i
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = maxincticker
        ws.Cells(2, 17).Value = maxinc
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = maxdecticker
        ws.Cells(3, 17).Value = maxdec
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = maxvolticker
        ws.Cells(4, 17).Value = maxvol
        
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("A:Q").AutoFit

    Next ws
End Sub