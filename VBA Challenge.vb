


Sub Loopthroughworksheets()


Dim ws  As Worksheet

For Each ws In ThisWorkbook.Worksheets

Dim x As Long
Dim J As Integer
Dim opening As Double
Dim closing As Double
Dim Change As Double
Dim percent As Double
Dim Ticker As String
Dim stockvolume As Double
Dim lastrow As Long
Dim Greatest_increase As Double
Dim Greatest_increaseTicker As String
Dim Greatest_decrease As Double
Dim Greatest_decreaseTicker As String
Dim Greatest_volume As Double

Greatest_increase = 0
Greatest_decrease = 0
Greatest_volume = 0



'Determine last row
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row



'Add Columns
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
ws.Cells(2, 16).Value = "Greatest % increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"



opening = ws.Cells(2, 3).Value
J = 2
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
For x = 2 To lastrow
'Find the place where the stocks change
'If the ticker changes then do something
Volume = Volume + ws.Cells(x, 7).Value
 If Cells(x, 1).Value <> Cells(x + 1, 1).Value Then
  Ticker = ws.Cells(x, 1).Value
  closing = ws.Cells(x, 6).Value
  Change = closing - opening
  'Find the change by subtracting closing -opening
  percent_change = (closing - opening) / opening
  ws.Cells(J, 10).Value = Ticker
  ws.Cells(J, 11).Value = Change
   If (Cells(J, 11).Value > 0) Then
     ws.Cells(J, 11).Interior.ColorIndex = 4
    ElseIf (Cells(J, 11).Value < 0) Then
    ws.Cells(J, 11).Interior.ColorIndex = 3
    End If
  ws.Cells(J, 12).Value = percent_change
    ws.Cells(J, 12).NumberFormat = "0.00%"
   If (Greatest_increase < percent_change) Then
   Greatest_increase = percent_change
   Greatest_increaseTicker = Ticker
   End If
   If (Greatest_decrease > percent_change) Then
   Greatest_decrease = percent_change
   Greatest_decreaseTicker = Ticker
    End If
   
   ws.Cells(J, 13).Value = Volume
   If (Greatest_volume < Volume) Then
   Greatest_volume = Volume
   Greatest_volumeTicker = Ticker

   
   End If
   
   'store the volume in the greatest volume variable
   
   
  J = J + 1
 opening = Cells(x + 1, 3).Value

Volume = 0


End If


Next x

ws.Cells(2, 17).Value = Greatest_increaseTicker
ws.Cells(2, 18).Value = Greatest_increase
ws.Cells(2, 18).NumberFormat = "0.00%"
ws.Cells(3, 17).Value = Greatest_decreaseTicker
ws.Cells(3, 18).Value = Greatest_decrease
ws.Cells(3, 18).NumberFormat = "0.00%"
ws.Cells(4, 17).Value = Greatest_volumeTicker
ws.Cells(4, 18).Value = Greatest_volume



Next ws

End Sub
