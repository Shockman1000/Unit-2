VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub challenge1():
    Dim ws As Worksheet
    Dim total As Double
    Dim row As Integer

    For Each ws In Worksheets
        total = 0
        row = 0

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).row
        
        For i = 2 To RowCount
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                total = total + ws.Cells(i, 7).Value
                ws.Range("I" & 2 + row).Value = ws.Cells(i, 1).Value

                ws.Range("J" & 2 + row).Value = total
                total = 0
                row = row + 1
            Else
                total = total + ws.Cells(i, 7).Value
            End If
        Next i
        total = 0
        row = 0
    Next ws
End Sub


