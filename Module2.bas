Attribute VB_Name = "Module2"
Sub InitializeAccounts()
    Dim wsKontoplan As Worksheet
    Dim accountRow As Long
    Dim lastRow As Long
    Dim kontoNummer As String
    Dim benamning As String
    Dim saldo As Double
    
    ' Set reference to Kontoplan sheet
    Set wsKontoplan = ThisWorkbook.Sheets("Kontoplan")
    
    ' Get the last row in Kontoplan sheet
    lastRow = wsKontoplan.Cells(wsKontoplan.Rows.Count, "G").End(xlUp).Row
    
    ' Loop through each row in Kontoplan to create account sheets
    For accountRow = 2 To lastRow ' Assuming the first row is headers
        If wsKontoplan.Cells(accountRow, "J").Value = "TRUE" Then ' Check if account is active
            kontoNummer = wsKontoplan.Cells(accountRow, "G").Value
            benamning = wsKontoplan.Cells(accountRow, "H").Value
            saldo = wsKontoplan.Cells(accountRow, "K").Value
            CreateAccountSheet kontoNummer, benamning, saldo
        End If
    Next accountRow
End Sub

Sub CreateAccountSheet(accountNumber As String, benamning As String, startingBalance As Double)
    Dim newSheet As Worksheet
    
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sets(accountNumber)
    On Error GoTo 0
    
    If newSheet Is Nothing Then
        ' Create a new sheet
        Set newSheet = ThisWorkbook.Sheets.Add
        newSheet.Name = accountNumber
        
        ' Set up headers and starting balance
        With newSheet
            .Range("A1").Value = "Konto"
            .Range("B1").Value = "Benämning"
            .Range("C1").Value = "Verifikationsnummer"
            .Range("D1").Value = "Datum"
            .Range("E1").Value = "Kostnadställe"
            .Range("F1").Value = "Projekt"
            .Range("G1").Value = "Verifikationstext"
            .Range("H1").Value = "Transaktionstextext"
            .Range("J1").Value = "Debet"
            .Range("K1").Value = "Kredit"
            .Range("L1").Value = "Saldo"
            .Range("A2").Value = accountNumber
            .Range("B2").Value = benamning
            .Range("J2").Value = startingBalance ' Starting saldo
        End With
    End If
End Sub

