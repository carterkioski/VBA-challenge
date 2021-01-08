Sub stocks():
    For Each ws In Worksheets
        Dim volume As LongLong
        Dim max_volume As LongLong
        Dim first As Double
        Dim last As Double
        Dim change As Double
        Dim max_gain As Double
        Dim max_loss As Double
        Dim output As Integer
        Dim ticker As String
        Dim max_gain_ticker As String
        Dim max_loss_ticker As String
        Dim max_volume_ticker As String
        output = 2
        volume = 0
        max_volume = 0
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock"
        max_gain = 0
        max_loss = 0
        ticker = ws.Cells(2, 1).Value
        first = ws.Cells(2, 3).Value
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row - 1
            'increment volume
            If first = 0 Then
                first = ws.Cells(i, 3).Value
            End If
            
            volume = volume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'print values
                last = ws.Cells(i, 6).Value
                change = last - first
                If Not first = 0 Then
                    If change / first < max_loss Then
                        max_loss = change / first
                        max_loss_ticker = ticker
                    ElseIf change / first > max_gain Then
                        max_gain = change / first
                        max_gain_ticker = ticker
                    End If
                End If
                If volume > max_volume Then
                    max_volume = volume
                    max_volume_ticker = ticker
                End If
                
                ws.Cells(output, 9).Value = ticker
                ws.Cells(output, 10).Value = change
                If Not (change = 0 And first = 0) Then
                    ws.Cells(output, 11) = Format(change / first, "percent")
                End If
                ws.Cells(output, 12).Value = volume
                'reset values
                ticker = ws.Cells(i + 1, 1)
                first = ws.Cells(i + 1, 3).Value
                volume = 0
                If ws.Cells(output, 10) < 0 Then
                    ws.Cells(output, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(output, 10).Interior.ColorIndex = 4
                End If
                output = output + 1
            Else
            End If
        Next i
        ws.Cells(1, 14).Value = "Max Gain Ticker"
        ws.Cells(2, 14).Value = max_gain_ticker
        ws.Cells(1, 15).Value = "Max Gain"
        ws.Cells(2, 15).Value = Format(max_gain, "percent")
        ws.Cells(1, 16).Value = "Max Loss Ticker"
        ws.Cells(2, 16).Value = max_loss_ticker
        ws.Cells(1, 17).Value = "Max Loss"
        ws.Cells(2, 17).Value = Format(max_loss, "percent")
        ws.Cells(1, 18).Value = "Max Volume Ticker"
        ws.Cells(2, 18).Value = max_volume_ticker
        ws.Cells(1, 19).Value = "Max Volume"
        ws.Cells(2, 19).Value = max_volume
        
    Next ws
End Sub
