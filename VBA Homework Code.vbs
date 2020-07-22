Attribute VB_Name = "Module1"
Sub free()

'Loop Worksheet
For Each ws In Worksheets
'Define variables
Dim openStock As Double
Dim closeStock As Double
Dim Ticker As String
Dim OutputRow As String
OutputRow = 2
Dim vol As Double

'Adding header
ws.Cells(1, 9).Value = "Ticker Symbol"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Finding last row
Dim lastRow As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Ticker = ws.Cells(2, 1).Value
openStock = ws.Cells(2, 3).Value

'Looping
For i = 2 To lastRow
    If (Ticker <> ws.Cells(i, 1).Value) Then
        closeStock = ws.Cells(i - 1, 6).Value
        ws.Cells(OutputRow, 10).Value = closeStock - openStock
        If openStock <> 0 Then
        ws.Cells(OutputRow, 11).Value = (closeStock - openStock) / openStock
        Else
        ws.Cells(OutputRow, 11).Value = "Cannot be calculated"
        End If
    
        ws.Cells(OutputRow, 9).Value = Ticker
        Ticker = ws.Cells(i, 1).Value
        
        ws.Cells(OutputRow, 12).Value = vol
        vol = 0
        openStock = ws.Cells(i, 3).Value
        OutputRow = OutputRow + 1
        
    Else
    vol = ws.Cells(i, 7).Value + vol
    
       
    End If
        
Next i

'Finding last row of percent change
Dim yc_lastRow As Double
    yc_lastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
'Loop conditional formatting
For i = 2 To yc_lastRow
    If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i




Next ws
End Sub

