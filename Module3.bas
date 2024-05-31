Attribute VB_Name = "Module3"
Sub TaBortAllaHuvudboksflikar()
    Dim ws As Worksheet
    Dim wsKontoplan As Worksheet
    Dim kontoNummer As String
    Dim lastRow As Long
    Dim i As Long
    
    Set wsKontoplan = ThisWorkbook.Sheets("Kontoplan")
    lastRow = wsKontoplan.Cells(wsKontoplan.Rows.Count, "G").End(xlUp).Row
    
    For i = 2 To lastRow
        kontoNummer = wsKontoplan.Cells(i, "G").Value
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(kontoNummer)
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
        On Error GoTo 0
    Next i
End Sub

