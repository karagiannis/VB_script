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

'Funktionen Worksheet_Change anv�nds f�r att automatiskt
'hantera �ndringar som g�rs i specificerade kolumner p� arbetsbladet Sheet5.
'N�r en anv�ndare trycker p� Enter efter att ha �ndrat en cell, kommer denna
'funktion att avg�ra om �ndringen gjordes i kolumnen f�r kontonummer, debet eller kredit.
'Baserat p� detta kommer motsvarande hanteringsfunktion att k�ras f�r att
'uppdatera relaterade data och ber�kningar.

    On Error GoTo ErrorHandler ' Aktivera felhantering
    Application.EnableEvents = False ' Inaktivera events f�r att undvika loopar

    ' Hantera �ndringar i kolumn Konto (kolumn 1 i ColumnNumbers)
    If Not Intersect(Target, Me.Range(Me.Cells(2, ColumnNumbers.Konto), Me.Cells(Me.Rows.Count, ColumnNumbers.Konto))) Is Nothing Then
        HandleAccountChange Target
    End If
    
    ' Hantera �ndringar i kolumn Debet eller Kredit (kolumn 12 och 13 i ColumnNumbers)
    If Not Intersect(Target, Me.Range(Me.Cells(2, ColumnNumbers.debet), Me.Cells(Me.Rows.Count, ColumnNumbers.debet))) Is Nothing Or _
       Not Intersect(Target, Me.Range(Me.Cells(2, ColumnNumbers.kredit), Me.Cells(Me.Rows.Count, ColumnNumbers.kredit))) Is Nothing Then
        HandleDebetKreditChange Target
        UpdateDiff ' Uppdatera diffen efter varje �ndring i debet eller kredit
    End If

    Application.EnableEvents = True ' �teraktivera events
    Exit Sub

ErrorHandler:
    ' Hantera fel
    MsgBox "Ett fel intr�ffade: " & Err.Description, vbExclamation
    Application.EnableEvents = True ' Se till att events alltid �teraktiveras
End Sub
Private Sub HandleAccountChange(ByVal Target As Range)

'Funktionen HandleAccountChange hanterar h�ndelser i den f�rsta kolumnen
'N�r en anv�ndare anger ett kontonummer kontrollerar funktionen om kontot finns i kontoplanen.
'Den h�mtar ocks� tillh�rande ben�mning och initialt saldo, och uppdaterar dessa v�rden i bokf�ringsbladet.
'Om kontot redan finns i tillf�llighetsytan summerar funktionen alla relaterade v�rden.

    Dim wsKontoplan As Worksheet
    Dim kontoNummer As String
    Dim benamning As String
    Dim saldo As Double
    Dim accountRow As Long
    Dim tempRow As Long
    Dim tempCol As Long
    Dim currentSaldo As Double
    
    ' Set reference to Kontoplan sheet
    Set wsKontoplan = ThisWorkbook.Sheets("Kontoplan")
    
    For Each cell In Target
        If cell.Column = ColumnNumbers.Konto Then
            kontoNummer = cell.Value
            If IsNumeric(kontoNummer) Then ' Verifiera att kontonumret �r et numeriskt v�rde
                benamning = GetBenamning(wsKontoplan, kontoNummer) 'H�mta ben�mningen enligt BAS-kontoplanen
                If benamning <> "" Then
                    Me.Cells(cell.Row, ColumnNumbers.Ben�mning).Value = benamning
                Else
                   'Kontonumret finns inte i BAS-kontoplanen
                    Me.Cells(cell.Row, ColumnNumbers.Ben�mning).Value = ""
                    MsgBox "Konto " & kontoNummer & " hittades inte i kontoplanen.", vbExclamation
                    Exit Sub
                End If
                
                ' Kontrollera om kontot redan finns i tillf�llighetsytan d� har tidigare bokf�ringsrad �ndrat
                'ing�ende balansen redan
                tempRow = GetAccountRowFromTillf�llighetsytan(Me, kontoNummer) 'Raden i tillf�llighetsytan d�r kontonumret �r skrivet
                If tempRow > 0 Then
                    ' Konto finns redan i tillf�llighetsytan, summera alla v�rden fr�n kolumn Ber�kningar och till h�ger
                    'initialsaldot innan aktuella rads debet/kredit-operationer
                    saldo = GetSaldoFromTillf�llighetsytan(ColumnNumbers.ber�kningar, CLng(kontoNummer))
                    
                     ' Kontrollera att tillf�llighetsytan �r korrekt
                    Dim previousEndSaldo As Double
                    previousEndSaldo = GetPreviousEndSaldo(kontoNummer, cell.Row)
                    If saldo <> previousEndSaldo Then
                        MsgBox "Tillf�llighetsytan �r korrumperad. Kontakta support!", vbExclamation
                        Exit Sub
                    End If
                Else
                    ' Konto finns inte i tillf�llighetsytan, kolla om kontot �r aktiverat
                    accountRow = GetAccountRow(wsKontoplan, kontoNummer)
                    If accountRow > 0 Then
                        ' Kontrollera om kontot �r aktiverat
                        If wsKontoplan.Cells(accountRow, "J").Value = "aktiverad" Then
                            ' Konto �r aktiverat, h�mta startsaldo fr�n huvudboken
                            saldo = GetSaldoFromHuvudbok(kontoNummer)
                        Else
                            ' Konto �r inte aktiverat, fr�ga anv�ndaren om de vill aktivera det
                            If MsgBox("Kontot �r inte aktiverat. Vill du aktivera kontot?", vbYesNo + vbQuestion, "Aktivera konto") = vbYes Then
                                saldo = 0
                                wsKontoplan.Cells(accountRow, "J").Value = "aktiverad"
                            Else
                                Me.Cells(cell.Row, ColumnNumbers.Ben�mning).Value = ""
                                saldo = ""
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox "Konto " & kontoNummer & " hittades inte i kontoplanen.", vbExclamation
                        Exit Sub
                    End If
                    
                    ' Skriv kontonummer och saldo i tillf�llighetsytan
                    tempRow = Me.Cells(Me.Rows.Count, ColumnNumbers.kontof�r�ndringar).End(xlUp).Row + 1
                    If tempRow < 2 Then tempRow = 2
                    Me.Cells(tempRow, ColumnNumbers.kontof�r�ndringar).Value = kontoNummer
                    Me.Cells(tempRow, ColumnNumbers.ber�kningar).Value = saldo
                End If
                
                ' Skriv initialt saldo i kolumn Saldo
                Me.Cells(cell.Row, ColumnNumbers.saldo).Value = saldo
            End If
        End If
    Next cell
End Sub


Private Sub HandleDebetKreditChange(ByVal Target As Range)

'Denna funktion hanterar h�ndelser i debet- och kredit-kolumnerna

    Dim bokf�ringsradNummer As Long
    Dim kontoNummer As String
    Dim initialSaldo As Double
    Dim debetBelopp As Double
    Dim kreditBelopp As Double
    Dim tempRow As Long
    Dim tempCol As Long
    Dim currentSaldo As Double
    
    For Each cell In Target
        If cell.Column = ColumnNumbers.debet Or cell.Column = ColumnNumbers.kredit Then ' Om �ndringen �r i kolumn D (Debet) eller E (Kredit)
           bokf�ringsradNummer = cell.Row
            kontoNummer = Me.Cells(bokf�ringsradNummer, ColumnNumbers.Konto).Value ' H�mta kontonumret fr�n kolumn A p� samma rad
            
            ' Debug meddelande f�r att kontrollera v�rdet och typen av kontoNummer
            Debug.Print "kontoNummer: " & kontoNummer & " (Typ: " & TypeName(kontoNummer) & ")"
            
            ' If IsNumeric(kontoNummer) And kontoNummer <> "" Then
            If kontoNummer <> "" Then
                ' H�mta saldot fr�n tillf�llighetsytan
                tempRow = GetAccountRowFromTillf�llighetsytan(Me, kontoNummer)
                If tempRow > 0 Then
                    
                    ' H�mta debet och kredit belopp
                    If cell.Column = ColumnNumbers.debet Then
                        debetBelopp = cell.Value
                        kreditBelopp = 0
                    Else
                        kreditBelopp = cell.Value
                        debetBelopp = 0
                    End If
                    
                    ' Uppdatera saldot i tillf�llighetsytan
                    tempCol = Me.Cells(tempRow, Me.Columns.Count).End(xlToLeft).Column + 1
                    If tempCol < ColumnNumbers.ber�kningar + 1 Then tempCol = ColumnNumbers.ber�kningar + 1
                    Me.Cells(tempRow, tempCol).Value = debetBelopp - kreditBelopp
                    
                    ' Summera alla v�rden fr�n kolumn K och till h�ger
                    currentSaldo = GetSaldoFromTillf�llighetsytan(ColumnNumbers.ber�kningar, kontoNummer)
                    
                    
                    ' Uppdatera saldot i ordinarie saldokolumn
                    Me.Cells(bokf�ringsradNummer, ColumnNumbers.saldo).Value = currentSaldo
                    
                    ' L�gg till ta-bort-rad-knapp om den inte redan finns
                    If Me.Cells(bokf�ringsradNummer, ColumnNumbers.diff).Value = "" Then
                        AddDeleteButton bokf�ringsradNummer
                    End If
                Else
                    MsgBox "Konto " & kontoNummer & " hittades inte i tillf�llighetsytan.", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    Next cell
End Sub
Function columnLetter(columnNumber As Integer) As String
    Dim columnBokst�ver As String
    columnBokst�ver = ""
    
    Do While columnNumber > 0
        columnBokst�ver = Chr((columnNumber - 1) Mod 26 + 65) & columnLetter
        columnNumber = (columnNumber - 1) \ 26
    Loop
    
    columnLetter = columnBokst�ver
End Function
Private Sub UpdateDiff()
    Dim lastRow As Long
    Dim diffRow As Long
    Dim debetSum As Double
    Dim kreditSum As Double
    Dim debetColLetter As String
    Dim kreditColLetter As String
    
    ' Best�m sista ifyllda raden i kolumn 1 (antagande att kolumn 1 alltid har data)
    lastRow = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row
    diffRow = lastRow + 10
    
    ' H�mta kolumnbokst�ver f�r Debet och Kredit
    debetColLetter = columnLetter(ColumnNumbers.debet)
    kreditColLetter = columnLetter(ColumnNumbers.kredit)
    
    ' Anv�nd kolumnbokst�verna i Range-funktionen f�r att summera debet och kredit
    debetSum = Application.WorksheetFunction.Sum(Me.Range(debetColLetter & "2:" & debetColLetter & lastRow))
    kreditSum = Application.WorksheetFunction.Sum(Me.Range(kreditColLetter & "2:" & kreditColLetter & lastRow))
    
    ' Ta bort tidigare diff-ber�kningar fr�n rad 9 under sista ifyllda raden
    Me.Rows(lastRow + 9).ClearContents
    
    ' L�gg till ny diff-ber�kning p� rad 10 under sista ifyllda raden
    Me.Cells(diffRow, ColumnNumbers.diff - 1).Value = "Diff"
    Me.Cells(diffRow, ColumnNumbers.diff).Value = debetSum - kreditSum
End Sub

' H�mta ben�mningen f�r ett konto
Private Function GetBenamning(ws As Worksheet, kontoNummer As String) As String
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow ' F�ruts�tter att den f�rsta raden �r headers
        If ws.Cells(i, "G").Value = kontoNummer Then
            GetBenamning = ws.Cells(i, "H").Value
            Exit Function
        End If
    Next i
    GetBenamning = ""
End Function

' H�mta raden f�r ett konto
Private Function GetAccountRow(ws As Worksheet, kontoNummer As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow ' F�ruts�tter att den f�rsta raden �r headers
        If ws.Cells(i, "G").Value = kontoNummer Then
            GetAccountRow = i
            Exit Function
        End If
    Next i
    GetAccountRow = 0
End Function

' H�mta raden f�r ett konto i tillf�llighetsytan
Private Function GetAccountRowFromTillf�llighetsytan(ws As Worksheet, kontoNummer As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, ColumnNumbers.kontof�r�ndringar).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow ' F�ruts�tter att tillf�llighetsytan b�rjar p� rad 2
        If ws.Cells(i, ColumnNumbers.kontof�r�ndringar).Value = kontoNummer Then
            GetAccountRowFromTillf�llighetsytan = i
            Exit Function
        End If
    Next i
    GetAccountRowFromTillf�llighetsytan = 0
End Function
Private Function GetSaldoFromHuvudbok(kontoNummer As String) As Double
    Dim saldo As Double
    Dim lastRow As Long
    Dim wsAccount As Worksheet

    On Error Resume Next
    Set wsAccount = ThisWorkbook.Sheets(kontoNummer)
    On Error GoTo 0

    If Not wsAccount Is Nothing Then
        lastRow = wsAccount.Cells(wsAccount.Rows.Count, 1).End(xlUp).Row
        saldo = wsAccount.Cells(lastRow, ColumnNumbers.saldo).Value
    Else
        saldo = 0
    End If

    GetSaldoFromHuvudbok = saldo
End Function

Function GetSaldoFromTillf�llighetsytan(startCol As Long, kontoNummer As String) As Double
    Dim tempRow As Long
    Dim tempCol As Long
    Dim saldo As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Bokf�ring")
    
    ' Hitta raden f�r kontonumret i tillf�llighetsytan
    tempRow = GetAccountRowFromTillf�llighetsytan(ws, kontoNummer)
    
    ' Kontrollera om raden hittades
    If tempRow > 0 Then
        saldo = ws.Cells(tempRow, startCol).Value
        tempCol = startCol + 1
        
        ' Summera alla v�rden fr�n startCol och till h�ger
        Do While ws.Cells(tempRow, tempCol).Value <> ""
            saldo = saldo + ws.Cells(tempRow, tempCol).Value
            tempCol = tempCol + 1
        Loop
    Else
        saldo = 0 ' Om kontonumret inte hittades, s�tt saldo till 0
    End If
    
    GetSaldoFromTillf�llighetsytan = saldo
End Function

Function GetPreviousEndSaldo(kontoNummer As String, currentRow As Long) As Double
    Dim wsBokforing As Worksheet
    Dim saldo As Double
    Dim i As Long
    
    Set wsBokforing = ThisWorkbook.Sheets("Bokf�ring")
    
    saldo = 0
    For i = currentRow - 1 To 2 Step -1
        If wsBokforing.Cells(i, ColumnNumbers.Konto).Value = kontoNummer Then
            saldo = wsBokforing.Cells(i, ColumnNumbers.saldo).Value
            Exit For
        End If
    Next i
    
    GetPreviousEndSaldo = saldo
End Function

Sub AddDeleteButton(rowNumber As Long)
    Dim ws As Worksheet
    Dim btn As Shape
    Dim btnName As String
    Dim btnCaption As String

    Set ws = ThisWorkbook.Sheets("Bokf�ring")

    ' Definiera namn och titel f�r knappen
    btnName = "DeleteButton_" & rowNumber
    btnCaption = "Delete"

    ' Kontrollera om knappen redan finns p� raden
    On Error Resume Next
    Set btn = ws.Shapes(btnName)
    On Error GoTo 0

    If btn Is Nothing Then
        ' Skapa ny knapp
        Set btn = ws.Shapes.AddFormControl(xlButtonControl, _
                                          ws.Cells(rowNumber, ColumnNumbers.diff).Left, _
                                          ws.Cells(rowNumber, ColumnNumbers.diff).Top, _
                                          ws.Cells(rowNumber, ColumnNumbers.diff).Width, _
                                          ws.Cells(rowNumber, ColumnNumbers.diff).Height)

        ' St�ll in knappens egenskaper
        With btn
            .Name = btnName
            .OnAction = "DeleteRow"
            .TextFrame.Characters.Text = btnCaption
        End With
    End If
End Sub


Public Sub UpdateTillf�llighetsytanEfterBorttagAvRad(kontoNummer As String)
    Dim ws As Worksheet
    Dim tempRow As Long
    Dim lastTempCol As Long

    Set ws = ThisWorkbook.Sheets("Bokf�ring")
    
    ' Hitta raden f�r kontonumret i tillf�llighetsytan
    tempRow = GetAccountRowFromTillf�llighetsytan(ws, kontoNummer)
    Debug.Print "tempRow f�r konto " & kontoNummer & ": " & tempRow
    
    If tempRow > 0 Then
        ' Hitta sista kolumnen med ett v�rde i tillf�llighetsytan
        lastTempCol = ws.Cells(tempRow, ws.Columns.Count).End(xlToLeft).Column
        Debug.Print "lastTempCol: " & lastTempCol
        
        ' Kolla om vi �r i ber�kningsomr�det
        If lastTempCol >= ColumnNumbers.ber�kningar Then
            ' L�gg till fels�kning f�r att se vad som orsakar problemet
            Debug.Print "tempRow: " & tempRow
            Debug.Print "lastTempCol: " & lastTempCol
            Debug.Print "Cell Value: " & ws.Cells(tempRow, lastTempCol).Value
            
            ' Ta bort sista v�rdet i tillf�llighetsytan
            ws.Cells(tempRow, lastTempCol).Value = ""
        End If
    Else
        ' Om raden inte hittas, skriv ut ett meddelande f�r fels�kning
        Debug.Print "Kontonumret " & kontoNummer & " hittades inte i tillf�llighetsytan."
    End If
End Sub


