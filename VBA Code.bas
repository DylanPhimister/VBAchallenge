Attribute VB_Name = "Module1"
Sub Challenge()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Ticker As String
Dim QuarterlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
TotalVolume = 0
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim i As Long
Dim Column As Integer
Column = 1
Dim Row As Double
Row = 2


Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Quarterly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Volume"

OpenPrice = ws.Cells(2, Column + 2).Value


For i = 2 To LastRow

Range("K2:K" & LastRow).NumberFormat = "0.00%"


If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        Ticker = Cells(i, Column).Value
        Cells(Row, Column + 8).Value = Ticker
        
    ClosePrice = ws.Cells(i, Column + 5).Value
    
   QuarterlyChange = ClosePrice - OpenPrice
        Cells(Row, Column + 9).Value = QuarterlyChange
        
    If ClosePrice = 0 And OpenPrice = 0 Then
    PercentChange = 0
    ElseIf ClosePrice <> 0 And OpenPrice = 0 Then
    PercentChange = 1
    Else
    PercentChange = QuarterlyChange / OpenPrice
        Cells(Row, Column + 10).Value = PercentChange
    End If
    
    OpenPrice = Cells(i + 1, Column + 2)
    TotalVolume = TotalVolume + Cells(i, Column + 6).Value
    Cells(Row, Column + 11).Value = TotalVolume
    Row = Row + 1
    OpenPrice = Cells(i + 1, Column + 2)
    TotalVolume = 0
    
    Else
    TotalVolume = TotalVolume + Cells(i, Column + 6).Value
    
End If

Next i

LastRowQuarterlyChange = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row

For j = 2 To LastRowQuarterlyChange

If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
    Cells(j, Column + 9).Interior.ColorIndex = 4
    
    ElseIf Cells(j, Column + 9).Value < 0 Then
            Cells(j, Column + 9).Interior.ColorIndex = 3
            
End If

Next j

    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    For x = 2 To LastRowQuarterlyChange
    
    If Cells(x, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRowQuarterlyChange)) Then
        Cells(2, 16).Value = Cells(x, Column + 8).Value
        Cells(2, 17).Value = Cells(x, Column + 10).Value
        Cells(2, 17).NumberFormat = "0.00%"
    
    ElseIf Cells(x, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRowQuarterlyChange)) Then
        Cells(3, 16).Value = Cells(x, Column + 8).Value
        Cells(3, 17).Value = Cells(x, Column + 10).Value
        Cells(3, 17).NumberFormat = "0.00%"
    
    ElseIf Cells(x, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowQuarterlyChange)) Then
        Cells(4, 16).Value = Cells(x, Column + 8).Value
        Cells(4, 17).Value = Cells(x, Column + 11).Value
End If

Next x

Next ws

End Sub


