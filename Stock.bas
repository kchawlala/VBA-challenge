Attribute VB_Name = "Module1"

Sub WallStreet()
For Each ws In Worksheets
Dim ticker As String
Dim volume As Double
Dim openYear As Double
Dim closeYear As Double
Dim Yearly_change As Double
Dim Yearly_percentage As Double
Dim row As Long
volume = 0
ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 10).Value = "Yearly_change"
ws.Cells(1, 11).Value = "Yearly_percentage"
ws.Cells(1, 12).Value = "Total Stock Volume"
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
row = 2
For i = 2 To LastRow
    If openYear = 0 Then
        openYear = ws.Cells(i, 3).Value
    End If
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        closeYear = ws.Cells(i, 6).Value
        Yearly_change = closeYear - openYear
        ticker = ws.Cells(i, 1).Value
        volume = volume + ws.Cells(i, 7).Value
        
        If Yearly_change <> 0 Then
           Yearly_percentage = Yearly_change / openYear
        Else
           Yearly_percentage = 0
        End If
        
        ws.Range("J" & row).Value = Yearly_change
        ws.Range("I" & row).Value = ticker
        ws.Range("K" & row).Value = Yearly_percentage
        ws.Range("L" & row).Value = volume
        row = row + 1
        volume = 0
        openYear = 0
    Else
        volume = volume + ws.Cells(i, 7).Value
    End If
Next i
Next ws
End Sub












