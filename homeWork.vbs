Sub homeWork()

For Each ws In Worksheets
    ws.Cells(1, 10) = "TICKER"
    ws.Cells(1, 11) = "VOLUME"
    ws.Cells(1, 12) = "OPEN DAY 1"
    ws.Cells(1, 13) = "CLOSE LAST OF YEAR"
    ws.Cells(1, 14) = "CHANGE"
    ws.Cells(1, 15) = "% CHANGE"
    
    ws.Cells(1, 18) = "TICKER"
    ws.Cells(1, 19) = "VOLUME MAX"
    ws.Cells(1, 20) = "% MAX"
    ws.Cells(1, 21) = "% MIN"
    
   lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
   lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    ticketPointer = 2

    
    runningSum = ws.Cells(2, 7).Value
    openMax = ws.Cells(2, 3).Value
    
    maxOnPageVOLUME = -1
    maxOnPageVolumeTicker = ""
    
    maxOnPagePERCENTAGE = -1
    maxOnPagePERCENTAGETICKER = ""
    
      
    minOnPagePERCENTAGE = 100
    minOnPagePERCENTAGETICKER = ""
    
    
        For i = 2 To lastRow + 1
            If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
                runningSum = runningSum + ws.Cells(i + 1, 7).Value
        Else
             ws.Cells(ticketPointer, 10) = ws.Cells(i, 1).Value
             ws.Cells(ticketPointer, 11) = runningSum
             closeVaue = ws.Cells(i, 6).Value
             ws.Cells(ticketPointer, 12) = openMax
             ws.Cells(ticketPointer, 13) = closeVaue
             ws.Cells(ticketPointer, 14) = (openMax - closeVaue)
             
             If openMax = 0 Then
                openMax = 0.001
              End If
    
            ws.Cells(ticketPointer, 15).Value = ((openMax - closeVaue) * 100 / openMax)
                

             
             If ((openMax - closeVaue) > 0) Then
                ws.Cells(ticketPointer, 10).Interior.Color = RGB(0, 255, 0)
             Else
                ws.Cells(ticketPointer, 10).Interior.Color = RGB(255, 0, 0)
             End If
             
             
            openMax = ws.Cells(i, 3).Value
            runningSum = ws.Cells(i + 1, 7).Value
            ticketPointer = ticketPointer + 1
            
            If maxOnPageVOLUME < runningSum Then
                   maxOnPageVOLUME = runningSum
                   maxOnPageVolumeTicker = ws.Cells(i, 1).Value
            End If
                   
           If maxOnPagePERCENTAGE < ws.Cells(ticketPointer, 15).Value Then
                   maxOnPagePERCENTAGE = ws.Cells(ticketPointer, 15)
                   maxOnPagePERCENTAGETICKER = ws.Cells(i, 1).Value
            End If
            
            If minOnPagePERCENTAGE > ws.Cells(ticketPointer, 15).Value Then
                   minOnPagePERCENTAGE = ws.Cells(ticketPointer, 15)
                   minOnPagePERCENTAGETICKER = ws.Cells(i, 1).Value
            End If
            
        End If
   Next i
   
    ws.Cells(1, 18) = "TICKER"
    ws.Cells(1, 19) = "VOLUME MAX"
    ws.Cells(1, 20) = "% MAX"
    ws.Cells(1, 21) = "% MIN"
    
    ws.Cells(2, 18) = maxOnPageVolumeTicker
    ws.Cells(2, 19) = maxOnPageVOLUME
    
    ws.Cells(3, 18) = maxOnPagePERCENTAGETICKER
    ws.Cells(3, 20) = maxOnPagePERCENTAGE
    
     ws.Cells(4, 18) = minOnPagePERCENTAGETICKER
     ws.Cells(4, 21) = minOnPagePERCENTAGE
    
   
Next ws

End Sub

