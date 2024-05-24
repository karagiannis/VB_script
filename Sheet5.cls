VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler ' Aktivera felhantering
    Application.EnableEvents = False ' Inaktivera events f�r att undvika loopar

    ' Hantera �ndringar i kolumn B (Konto)
    If Not Intersect(Target, Me.Range("B7:B" & Me.Rows.Count)) Is Nothing Then
        HandleAccountChange Target
    End If
    
    ' Anropa funktionen f�r att hantera debet och kredit �ndringar
    Debet_Kredit_Saldo_Change Target

    Application.EnableEvents = True ' �teraktivera events
    Exit Sub

ErrorHandler:
    ' Hantera fel
    MsgBox "Ett fel intr�ffade: " & Err.Description, vbExclamation
    Application.EnableEvents = True ' Se till att events alltid �teraktiveras
End Sub

Private Sub HandleAccountChange(ByVal Target As Range)
    Dim wsKontoplan As Worksheet
    Dim kontoNummer As Variant
    Dim ben�mning As Variant
    Dim saldo As Double
    Dim activateAccount As VbMsgBoxResult
    Dim accountRow As Long
    Dim accountActivated As Boolean
    
    ' Set reference to Kontoplan sheet
    Set wsKontoplan = ThisWorkbook.Sheets("Kontoplan")
    
    ' Loop through each changed cell
    For Each cell In Target
        If cell.Column = 2 Then
            kontoNummer = cell.Value
            If IsNumeric(kontoNummer) Then
                ben�mning = GetAccountName(kontoNummer, wsKontoplan)
                cell.Offset(0, 1).Value = ben�mning
                
                ' Find the row for the account number in Kontoplan sheet
                accountRow = Application.WorksheetFunction.Match(kontoNummer, wsKontoplan.Range("G:G"), 0)
                
                ' Check if the account is activated
                accountActivated = IsAccountActivated(accountRow, wsKontoplan)
                
                If accountActivated Then
                    ' Account is activated, get the starting balance from Kontoplan
                    saldo = wsKontoplan.Cells(accountRow, "K").Value
                Else
                    ' Account is not activated, ask the user if they want to activate it
                    activateAccount = MsgBox("Kontot �r inte aktiverat. Vill du aktivera kontot?", vbYesNo + vbQuestion, "Aktivera konto")
                    If activateAccount = vbYes Then
                        ' Activate the account with starting balance 0.0
                        saldo = 0#
                        wsKontoplan.Cells(accountRow, "J").Value = "TRUE"
                        CreateAccountSheet CStr(kontoNummer), saldo
                    Else
                        ' Clear the Ben�mning cell if the user does not want to activate the account
                        cell.Offset(0, 1).Value = ""
                        saldo = ""
                    End If
                End If
                
                ' Update the Saldo cell
                cell.Offset(0, 5).Value = saldo
            End If
        End If
    Next cell
End Sub
Private Function GetAccountName(kontoNummer As Variant, wsKontoplan As Worksheet) As Variant
    On Error Resume Next
    GetAccountName = Application.WorksheetFunction.VLookup(kontoNummer, wsKontoplan.Range("G:H"), 2, False)
    On Error GoTo 0
End Function
Private Sub Debet_Kredit_Saldo_Change(ByVal Target As Range)
    Dim kontoNummer As Variant
    Dim saldo As Double
    Dim rowSaldo As Double
    Dim cell As Range
    Dim wsAccount As Worksheet
    
    ' Kontrollera om �ndringen sker i kolumn E (Debet) eller F (Kredit) fr�n rad 8 och fram�t
    If Not Intersect(Target, Me.Range("E8:F" & Me.Rows.Count)) Is Nothing Then
        ' Loop through each changed cell
        For Each cell In Target
            If cell.Column = 5 Or cell.Column = 6 Then ' Kontrollera om �ndringen sker i kolumn E eller F
                kontoNummer = Me.Cells(cell.Row, "B").Value
                If IsNumeric(kontoNummer) Then
                    On Error Resume Next
                    Set wsAccount = ThisWorkbook.Sheets(CStr(kontoNummer))
                    On Error GoTo 0
                    If Not wsAccount Is Nothing Then
                        ' H�mta senaste saldo fr�n huvudboksfliken
                        rowSaldo = wsAccount.Cells(wsAccount.Rows.Count, "F").End(xlUp).Value
                        
                        ' Uppdatera saldo baserat p� debet/kredit
                        If IsNumeric(cell.Value) Then
                            If cell.Column = 5 Then ' Debet
                                rowSaldo = rowSaldo + cell.Value
                            ElseIf cell.Column = 6 Then ' Kredit
                                rowSaldo = rowSaldo - cell.Value
                            End If
                        End If
                        
                        ' Uppdatera Saldo i kolumn G
                        Me.Cells(cell.Row, "G").Value = rowSaldo
                        
                        ' L�gg till transaktionen i huvudboksfliken
                        Dim newRow As Long
                        newRow = wsAccount.Cells(wsAccount.Rows.Count, "A").End(xlUp).Row + 1
                        wsAccount.Cells(newRow, "A").Value = kontoNummer ' Konto
                        wsAccount.Cells(newRow, "B").Value = Me.Cells(cell.Row, "B").Value ' Ben�mning
                        wsAccount.Cells(newRow, "C").Value = Me.Cells(cell.Row, "A").Value ' Verifikationsnummer
                        wsAccount.Cells(newRow, "D").Value = Me.Range("G3").Value ' Registreringsdatum (fakturadatum)
                        wsAccount.Cells(newRow, "E").Value = Me.Cells(cell.Row, "F").Value ' Kostnadst�lle
                        wsAccount.Cells(newRow, "F").Value = Me.Cells(cell.Row, "G").Value ' Projekt
                        wsAccount.Cells(newRow, "G").Value = Me.Range("G4").Value ' Verifikationstext
                        wsAccount.Cells(newRow, "H").Value = cell.Offset(0, -2).Value ' Debet
                        wsAccount.Cells(newRow, "I").Value = cell.Offset(0, -1).Value ' Kredit
                        wsAccount.Cells(newRow, "J").Value = rowSaldo ' Uppdaterat
                        
                    End If
                End If
            End If
        Next cell
    End If
End Sub
