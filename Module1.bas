Attribute VB_Name = "Module1"
Sub allYearlyChange()
 
 Dim lastRow As Long
 
 Dim yearlyChange As Double
 Dim percentage As Double
 
 
 'Varables for the max/min table
 Dim minIncrease As Double
 Dim maxIncrease As Double
 Dim maxVol As Double
 Dim minTicker As String
 Dim maxTicker As String
 Dim volTicker As String
 
 
 'Loop through the sheets
 For Each ws In Worksheets
 
     'Find the last row
     lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     'Init the values
     openPrice = ws.Cells(2, 3).Value
     closePrice = ws.Cells(2, 6).Value
     volTotal = 0
     yearlyChange = closePrice - openPrice
     
     'Creating new columns
     ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 10).Value = "Yearly Change"
     ws.Cells(1, 11).Value = "Percentage Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
     rowNum = 2
     
     
     'Initialize the greatest increase/decrease/volume
     minIncrease = (closePrice - openPrice) / openPrice
     maxIncrease = (closePrice - openPrice) / openPrice
     maxVol = ws.Cells(2, 7).Value
     maxTicker = ws.Cells(2, 1).Value
     minTicker = ws.Cells(2, 1).Value
     volTicker = ws.Cells(2, 1).Value

 
    
     
      For i = 2 To lastRow - 1
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            closePrice = ws.Cells(i, 6).Value
            percentage = (closePrice - openPrice) / openPrice
            yearlyChange = closePrice - openPrice
            
            ws.Cells(rowNum, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(rowNum, 10).Value = yearlyChange
            
            'Conditional Formatting
            If yearlyChange < 0 Then
                ws.Cells(rowNum, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(rowNum, 10).Interior.ColorIndex = 4
            End If
            
            
            'Update max/min/vol
            If percentage < minIncrease Then
                 minIncrease = percentage
                 minTicker = ws.Cells(i, 1).Value
            End If
            If percentage > maxIncrease Then
                 maxIncrease = percentage
                 maxTicker = ws.Cells(i, 1).Value
            End If
            If volTotal > maxVol Then
                 maxVol = volTotal
                 volTicker = ws.Cells(i, 1).Value
            End If
            
            
            
            ws.Cells(rowNum, 11).Value = percentage
            ws.Cells(rowNum, 12).Value = volTotal
            
            rowNum = rowNum + 1
            volTotal = 0
            openPrice = ws.Cells(i + 1, 3).Value
            
            
        Else
            volTotal = volTotal + ws.Cells(i, 7).Value
            
            
        End If
     
     Next i
        
       'Calculate for the last row of the sheet
        volTotal = volTotal + ws.Cells(lastRow, 7).Value
        closePrice = ws.Cells(lastRow, 6).Value
        percentage = (closePrice - openPrice) / openPrice
        yearlyChange = closePrice - openPrice
            
        ws.Cells(rowNum, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(rowNum, 10).Value = yearlyChange
        'Conditional Formatting
         If yearlyChange < 0 Then
               ws.Cells(rowNum, 10).Interior.ColorIndex = 3
         Else
               ws.Cells(rowNum, 10).Interior.ColorIndex = 4
         End If
         
         ws.Cells(rowNum, 11).Value = percentage
         ws.Cells(rowNum, 12).Value = volTotal
         
         
    
         'Update the greatest increase/decrease/volume
        If percentage < minIncrease Then
         minIncrease = percentage
         minTicker = ws.Cells(lastRow, 1).Value
        End If
        If percentage > maxIncrease Then
            maxIncrease = percentage
            maxTicker = ws.Cells(lastRow, 1).Value
        End If
        If volTotal > maxVol Then
            maxVol = volTotal
            volTicker = ws.Cells(lastRow, 1).Value
        End If
    
    
    'Creating new columns for the new table
     ws.Cells(2, 15).Value = "Greatest % Increase"
     ws.Cells(3, 15).Value = "Greatest % Decrease"
     ws.Cells(4, 15).Value = "Greatest Total Volume"
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
     ws.Cells(2, 16).Value = maxTicker
     ws.Cells(3, 16).Value = minTicker
     ws.Cells(4, 16).Value = volTicker
     ws.Cells(2, 17).Value = maxIncrease
     ws.Cells(3, 17).Value = minIncrease
     ws.Cells(4, 17).Value = maxVol
     
     
            
    MsgBox ("Completed " + ws.Name)
  
  Next ws
    
End Sub




