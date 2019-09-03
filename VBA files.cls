VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub StockTracker()
For Each ws In Worksheets

Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Change As Double
Dim PChange As Double
Dim Volume As Double

Volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'counts the number of rows
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For R = 2 To lastrow
'Grab open Price
If ws.Cells(R - 1, 1).Value <> ws.Cells(R, 1).Value Then
OpenPrice = ws.Cells(R, 3).Value
End If
'Grab Ticker, Volume, ClosePrice
If ws.Cells(R + 1, 1).Value <> ws.Cells(R, 1).Value Then
Ticker = ws.Cells(R, 1).Value
Volume = Volume + ws.Cells(R, 7).Value
ClosePrice = ws.Cells(R, 6).Value
'Calculate Change and PChange
Change = ClosePrice - OpenPrice
If OpenPrice > 0 And ClosePrice > 0 Then
PChange = (Change / OpenPrice)
End If

'print stock ticker in summary table
ws.Range("I" & Summary_Table_Row).Value = Ticker
ws.Cells(1, 9).Value = "Ticker"

ws.Range("J" & Summary_Table_Row).Value = Change
ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
ws.Cells(1, 10).Value = "Yearly Change"

ws.Range("K" & Summary_Table_Row).Value = PChange
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
ws.Cells(1, 11).Value = "Percent Change"

ws.Range("L" & Summary_Table_Row).Value = Volume
ws.Cells(1, 12).Value = "Total Volume"

'Add one to summary table row
Summary_Table_Row = Summary_Table_Row + 1
'reset Volume
Volume = 0
Else
Volume = Volume + ws.Cells(R, 7).Value
End If

'Apply Color Pallette
If ws.Cells(R, 10).Value > 0 Then
ws.Cells(R, 10).Interior.ColorIndex = 4

ElseIf ws.Cells(R, 10).Value = 0 Then
ws.Cells(R, 10).Interior.ColorIndex = 2

Else
ws.Cells(R, 10).Interior.ColorIndex = 3

End If

Next R

Next ws

End Sub

