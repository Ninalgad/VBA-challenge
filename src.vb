Sub summariseTickers():
    Dim curVolume, pointer, i, lastRow As LongLong 
    Dim curTicker, ticker As String
    Dim curOpening, closing, yearlyChange, percentageChange As Double
    Dim ws As Worksheet


    For Each ws In ThisWorkbook.Worksheets
        curTicker = ws.Cells(2, 1).Value
        curOpening = ws.Cells(2, 3).Value
        curVolume = ws.Cells(2, 7).Value
        lastRow = ws.Range("A1").End(xlDown).Row
        pointer = 2
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volumes"
    
        For i = 3 To lastRow
            ticker = ws.Cells(i, 1).Value

            If (ticker <> curTicker) Or (i = lastRow) Then
                closing = ws.Cells(i - 1, 6).Value
                yearlyChange = closing - curOpening
                percentageChange = yearlyChange / curOpening
    
                'output
                ws.Cells(pointer, 9).Value = curTicker
                ws.Cells(pointer, 10).Value = yearlyChange
                ws.Cells(pointer, 12).Value = curVolume
                ws.Cells(pointer, 11).Value = FormatPercent(percentageChange) 

                colorIdx = 43
                If yearlyChange < 0 Then
                    colorIdx = 22
                End If
                ws.Cells(pointer, 10).Interior.ColorIndex = colorIdx
    
                'update
                curTicker = ticker
                curVolume = ws.Cells(i, 7).Value
                curOpening = ws.Cells(i, 3).Value
                pointer = pointer + 1
            Else
                curVolume = curVolume + ws.Cells(i, 7).Value
            End If
        Next i

        
        ' Maximial values functionality
        maxPctIncr = 0
        maxPctIncrTicker = ""
        minPctIncr = 0
        minPctIncrTicker = ""
        maxVol = 0
        maxVolTicker = ""
        
        For i = 2 To ws.Range("K1").End(xlDown).Row
            p = ws.Cells(i, 11).Value
            t = ws.Cells(i, 9).Value
            v = ws.Cells(i, 12).Value
            If p > maxPctIncr Then
                maxPctIncr = p
                maxPctIncrTicker = t
            ElseIf (p < minPctIncr) Then
                minPctIncr = p
                minPctIncrTicker = t
            End If
            
            If v > maxVol Then
                maxVol = v
                maxVolTicker = t
            End If
        Next i
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        
        ws.Range("P2").Value = maxPctIncrTicker
        ws.Range("P3").Value = minPctIncrTicker
        ws.Range("P4").Value = maxVolTicker
        
        ws.Range("Q2").Value = FormatPercent(maxPctIncr)
        ws.Range("Q3").Value = FormatPercent(minPctIncr)
        ws.Range("Q4").Value = maxVol
    Next ws
End Sub

          