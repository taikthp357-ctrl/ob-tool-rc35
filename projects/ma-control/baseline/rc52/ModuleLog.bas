Attribute VB_Name = "ModuleLog"

Option Explicit

Public Sub WriteLog(ByVal logType As String, ByVal content As String, Optional ByVal refInfo As String = "", Optional ByVal productCode As String = "")

Dim ws As Worksheet, r As Long

Set ws = ThisWorkbook.Worksheets("LOG")

r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1

If r < 6 Then r = 6

ws.Cells(r, 1).NumberFormat = "m/d/yy h:mm AM/PM"
ws.Cells(r, 1).value = Now
ws.Cells(r, 2).Resize(1, 5).Value2 = Array(ModuleMain.RunID, logType, refInfo, productCode, content)

With ws.Cells(r, 1).Resize(1, 6).Borders

 .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)

End With

End Sub

