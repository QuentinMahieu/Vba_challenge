Attribute VB_Name = "Module1"
Sub summary()

'set the header of the table
Dim ws As Worksheet
For Each ws In Worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"    'as the closing value - opening value
ws.Range("K1").Value = "Percent Change"  'yearly change/ closing value
ws.Range("L1").Value = "Total Stock"  'as the sum of volumes
ws.Range("N1").Value = "Ticker"
ws.Range("O1").Value = "Percent"
ws.Range("M2").Value = "Greatest % increase"
ws.Range("M3").Value = "Greatest % Decrease"
ws.Range("M4").Value = "Greatest % Decrease"

'find the last row of a worksheet
Dim lastRow As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim ticker As String
Dim yearlyOpen As Double
Dim yearlyClose As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalStock As Double
Dim cellCount As Double
Dim tableRow As Integer

tableRow = 1
totalStock = 0

'for looping
    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ticker = ws.Cells(i, 1).Value
        yearlyOpen = ws.Cells(i - cellCount, 3).Value
        yearlyClose = ws.Cells(i, 6).Value
        yearlyChange = yearlyClose - yearlyOpen
        
            If yearlyOpen <> 0 Then
                percentChange = yearlyChange / yearlyOpen
            Else:   percentChange = 0
            End If
            
        
        cellCount = 0
        tableRow = tableRow + 1
        
        totalStock = totalStock + ws.Cells(i, 7).Value
        ws.Cells(tableRow, 9).Value = ticker
        ws.Cells(tableRow, 12).Value = totalStock
        ws.Cells(tableRow, 10).Value = yearlyChange
        
            If yearlyChange = 0 Then
                ws.Cells(tableRow, 10).Interior.ColorIndex = 0
            ElseIf yearlyChange < 0 Then
                ws.Cells(tableRow, 10).Interior.ColorIndex = 3
            Else: ws.Cells(tableRow, 10).Interior.ColorIndex = 4
            End If
        
        ws.Cells(tableRow, 11).Value = percentChange
        
        totalStock = 0
        
        Else: totalStock = totalStock + ws.Cells(i, 7).Value
                cellCount = cellCount + 1
        End If
        
    Next i
        
    ws.Range("K2:K" & tableRow).NumberFormat = "0.00%"
    ws.Range("O2:O3").NumberFormat = "0.00%"
    ws.Range("O2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & tableRow))
    ws.Range("O3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & tableRow))
    ws.Range("O4").Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & tableRow))
    
       For i = 2 To tableRow
         If ws.Cells(i, 11).Value = ws.Range("O2").Value Then
            ws.Range("N2").Value = ws.Cells(i, 9).Value
        End If
         If ws.Cells(i, 11).Value = ws.Range("O3").Value Then
                ws.Range("N3").Value = ws.Cells(i, 9).Value
        End If
         If ws.Cells(i, 12).Value = ws.Range("O4").Value Then
                    ws.Range("N4").Value = ws.Cells(i, 9).Value
        End If
         
       Next i
        
Next ws

End Sub
