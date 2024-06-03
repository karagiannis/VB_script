Attribute VB_Name = "Module2"
Public Enum ColumnNumbers
    Konto = 1
    Benämning = 2
    Beskrivning = 3
    verifikationsserie = 4
    verNr = 5
    systemdatum = 6
    registreringsdatum = 7
    kostnadsställe = 8
    projekt = 9
    verifikationstext = 10
    transaktionsinfo = 11
    debet = 12
    kredit = 13
    saldo = 14
    diff = 15
    bokföringsunderlag = 16
    kontoförändringar = 17
    beräkningar = 18
    
End Enum



' Definiera strängkonstanter för headers
Public Const Konto_s As String = "Konto"
Public Const Benämning_s As String = "Benämning"
Public Const Beskrivning_s As String = "Beskrivning"
Public Const Verifikationsserie_s As String = "Verifikationsserie"
Public Const VerNr_s As String = "Ver.nr"
Public Const Systemdatum_s As String = "Systemdatum"
Public Const Registreringsdatum_s As String = "Registreringsdatum"
Public Const Kostnadsställe_s As String = "Kostnadsställe"
Public Const Projekt_s As String = "Projekt"
Public Const Verifikationstext_s As String = "Verifikationstext"
Public Const Transaktionsinfo_s As String = "Transaktionsinfo"
Public Const Debet_s As String = "Debet"
Public Const Kredit_s As String = "Kredit"
Public Const Saldo_s As String = "Saldo"
Public Const Diff_s As String = "Diff"
Public Const Bokföringsunderlag_s As String = "Bokföringsunderlag"
Public Const Kontoförändringar_s As String = "Kontoförändringar"
Public Const Beräkningar_s As String = "Beräkningar"


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
        If wsKontoplan.Cells(accountRow, "J").Value = "aktiverad" Then ' Check if account is active
            kontoNummer = wsKontoplan.Cells(accountRow, "G").Value
            benamning = wsKontoplan.Cells(accountRow, "H").Value
            saldo = wsKontoplan.Cells(accountRow, "K").Value
            CreateAccountSheet kontoNummer, benamning, saldo
        End If
    Next accountRow
    
    InitializeVerifikationslista
End Sub

Sub CreateAccountSheet(accountNumber As String, Benämning As String, startingBalance As Double)
    Dim newSheet As Worksheet
    
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets(accountNumber)
    On Error GoTo 0
    
    If newSheet Is Nothing Then
        ' Skapa ett nytt blad
        Set newSheet = ThisWorkbook.Sheets.Add
        newSheet.Name = accountNumber
        
        ' Initiera headers
        InitializeHeaders newSheet
        
        ' Sätt upp startsaldo
        With newSheet
            .Cells(2, ColumnNumbers.Konto).Value = accountNumber
            .Cells(2, ColumnNumbers.Benämning).Value = Benämning
            .Cells(2, ColumnNumbers.transaktionsinfo).Value = "Ingående Balans"
            .Cells(2, ColumnNumbers.saldo).Value = startingBalance ' Startsaldo
        End With
    End If
End Sub


Sub InitializeHeaders(ws As Worksheet)
    With ws
        .Cells(1, ColumnNumbers.Konto).Value = Konto_s
        .Cells(1, ColumnNumbers.Benämning).Value = Benämning_s
        .Cells(1, ColumnNumbers.Beskrivning).Value = Beskrivning_s
        .Cells(1, ColumnNumbers.verifikationsserie).Value = Verifikationsserie_s
        .Cells(1, ColumnNumbers.verNr).Value = VerNr_s
        .Cells(1, ColumnNumbers.systemdatum).Value = Systemdatum_s
        .Cells(1, ColumnNumbers.registreringsdatum).Value = Registreringsdatum_s
        .Cells(1, ColumnNumbers.kostnadsställe).Value = Kostnadsställe_s
        .Cells(1, ColumnNumbers.projekt).Value = Projekt_s
        .Cells(1, ColumnNumbers.verifikationstext).Value = Verifikationstext_s
        .Cells(1, ColumnNumbers.transaktionsinfo).Value = Transaktionsinfo_s
        .Cells(1, ColumnNumbers.debet).Value = Debet_s
        .Cells(1, ColumnNumbers.kredit).Value = Kredit_s
        .Cells(1, ColumnNumbers.saldo).Value = Saldo_s
        .Cells(1, ColumnNumbers.diff).Value = Diff_s
        .Cells(1, ColumnNumbers.bokföringsunderlag).Value = Bokföringsunderlag_s
        .Cells(1, ColumnNumbers.kontoförändringar).Value = Kontoförändringar_s
        .Cells(1, ColumnNumbers.beräkningar).Value = Beräkningar_s
    End With
End Sub


Sub InitializeVerifikationslista()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Verifikationslista")
    
    ' Initiera headers
    InitializeHeaders ws
End Sub

Sub InitializeBokföring()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Bokföring")
    
    ' Initiera headers
    InitializeHeaders ws
End Sub

