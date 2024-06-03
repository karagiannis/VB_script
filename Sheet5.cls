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

'Funktionen Worksheet_Change används för att automatiskt
'hantera ändringar som görs i specificerade kolumner på arbetsbladet Sheet5.
'När en användare trycker på Enter efter att ha ändrat en cell, kommer denna
'funktion att avgöra om ändringen gjordes i kolumnen för kontonummer, debet eller kredit.
'Baserat på detta kommer motsvarande hanteringsfunktion att köras för att
'uppdatera relaterade data och beräkningar.

    On Error GoTo ErrorHandler ' Aktivera felhantering
    Application.EnableEvents = False ' Inaktivera events för att undvika loopar

    ' Hantera ändringar i kolumn Konto (kolumn 1 i ColumnNumbers)
    If Not Intersect(Target, Me.Range(Me.Cells(2, ColumnNumbers.Konto), Me.Cells(Me.Rows.Count, ColumnNumbers.Konto))) Is Nothing Then
        HandleAccountChange Target
    End If
    
    ' Hantera ändringar i kolumn Debet eller Kredit (kolumn 12 och 13 i ColumnNumbers)
    If Not Intersect(Target, Me.Range(Me.Cells(2, ColumnNumbers.debet), Me.Cells(Me.Rows.Count, ColumnNumbers.debet))) Is Nothing Or _
       Not Intersect(Target, Me.Range(Me.Cells(2, ColumnNumbers.kredit), Me.Cells(Me.Rows.Count, ColumnNumbers.kredit))) Is Nothing Then
        HandleDebetKreditChange Target
        UpdateDiff ' Uppdatera diffen efter varje ändring i debet eller kredit
    End If

    Application.EnableEvents = True ' Återaktivera events
    Exit Sub

ErrorHandler:
    ' Hantera fel
    MsgBox "Ett fel inträffade: " & Err.Description, vbExclamation
    Application.EnableEvents = True ' Se till att events alltid återaktiveras
End Sub
Private Sub HandleAccountChange(ByVal Target As Range)

'Funktionen HandleAccountChange hanterar händelser i den första kolumnen
'När en användare anger ett kontonummer kontrollerar funktionen om kontot finns i kontoplanen.
'Den hämtar också tillhörande benämning och initialt saldo, och uppdaterar dessa värden i bokföringsbladet.
'Om kontot redan finns i tillfällighetsytan summerar funktionen alla relaterade värden.

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
            If IsNumeric(kontoNummer) Then ' Verifiera att kontonumret är et numeriskt värde
                benamning = GetBenamning(wsKontoplan, kontoNummer) 'Hämta benämningen enligt BAS-kontoplanen
                If benamning <> "" Then
                    Me.Cells(cell.Row, ColumnNumbers.Benämning).Value = benamning
                Else
                   'Kontonumret finns inte i BAS-kontoplanen
                    Me.Cells(cell.Row, ColumnNumbers.Benämning).Value = ""
                    MsgBox "Konto " & kontoNummer & " hittades inte i kontoplanen.", vbExclamation
                    Exit Sub
                End If
                
                ' Kontrollera om kontot redan finns i tillfällighetsytan då har tidigare bokföringsrad ändrat
                'ingående balansen redan
                tempRow = GetAccountRowFromTillfällighetsytan(Me, kontoNummer) 'Raden i tillfällighetsytan där kontonumret är skrivet
                If tempRow > 0 Then
                    ' Konto finns redan i tillfällighetsytan, summera alla värden från kolumn Beräkningar och till höger
                    'initialsaldot innan aktuella rads debet/kredit-operationer
                    saldo = GetSaldoFromTillfällighetsytan(ColumnNumbers.beräkningar, CLng(kontoNummer))
                    
                     ' Kontrollera att tillfällighetsytan är korrekt
                    Dim previousEndSaldo As Double
                    previousEndSaldo = GetPreviousEndSaldo(kontoNummer, cell.Row)
                    If saldo <> previousEndSaldo Then
                        MsgBox "Tillfällighetsytan är korrumperad. Kontakta support!", vbExclamation
                        Exit Sub
                    End If
                Else
                    ' Konto finns inte i tillfällighetsytan, kolla om kontot är aktiverat
                    accountRow = GetAccountRow(wsKontoplan, kontoNummer)
                    If accountRow > 0 Then
                        ' Kontrollera om kontot är aktiverat
                        If wsKontoplan.Cells(accountRow, "J").Value = "aktiverad" Then
                            ' Konto är aktiverat, hämta startsaldo från huvudboken
                            saldo = GetSaldoFromHuvudbok(kontoNummer)
                        Else
                            ' Konto är inte aktiverat, fråga användaren om de vill aktivera det
                            If MsgBox("Kontot är inte aktiverat. Vill du aktivera kontot?", vbYesNo + vbQuestion, "Aktivera konto") = vbYes Then
                                saldo = 0
                                wsKontoplan.Cells(accountRow, "J").Value = "aktiverad"
                            Else
                                Me.Cells(cell.Row, ColumnNumbers.Benämning).Value = ""
                                saldo = ""
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox "Konto " & kontoNummer & " hittades inte i kontoplanen.", vbExclamation
                        Exit Sub
                    End If
                    
                    ' Skriv kontonummer och saldo i tillfällighetsytan
                    tempRow = Me.Cells(Me.Rows.Count, ColumnNumbers.kontoförändringar).End(xlUp).Row + 1
                    If tempRow < 2 Then tempRow = 2
                    Me.Cells(tempRow, ColumnNumbers.kontoförändringar).Value = kontoNummer
                    Me.Cells(tempRow, ColumnNumbers.beräkningar).Value = saldo
                End If
                
                ' Skriv initialt saldo i kolumn Saldo
                Me.Cells(cell.Row, ColumnNumbers.saldo).Value = saldo
            End If
        End If
    Next cell
End Sub


Private Sub HandleDebetKreditChange(ByVal Target As Range)

'Denna funktion hanterar händelser i debet- och kredit-kolumnerna

    Dim bokföringsradNummer As Long
    Dim kontoNummer As String
    Dim initialSaldo As Double
    Dim debetBelopp As Double
    Dim kreditBelopp As Double
    Dim tempRow As Long
    Dim tempCol As Long
    Dim currentSaldo As Double
    
    For Each cell In Target
        If cell.Column = ColumnNumbers.debet Or cell.Column = ColumnNumbers.kredit Then ' Om ändringen är i kolumn D (Debet) eller E (Kredit)
           bokföringsradNummer = cell.Row
            kontoNummer = Me.Cells(bokföringsradNummer, ColumnNumbers.Konto).Value ' Hämta kontonumret från kolumn A på samma rad
            
            ' Debug meddelande för att kontrollera värdet och typen av kontoNummer
            Debug.Print "kontoNummer: " & kontoNummer & " (Typ: " & TypeName(kontoNummer) & ")"
            
            ' If IsNumeric(kontoNummer) And kontoNummer <> "" Then
            If kontoNummer <> "" Then
                ' Hämta saldot från tillfällighetsytan
                tempRow = GetAccountRowFromTillfällighetsytan(Me, kontoNummer)
                If tempRow > 0 Then
                    
                    ' Hämta debet och kredit belopp
                    If cell.Column = ColumnNumbers.debet Then
                        debetBelopp = cell.Value
                        kreditBelopp = 0
                    Else
                        kreditBelopp = cell.Value
                        debetBelopp = 0
                    End If
                    
                    ' Uppdatera saldot i tillfällighetsytan
                    tempCol = Me.Cells(tempRow, Me.Columns.Count).End(xlToLeft).Column + 1
                    If tempCol < ColumnNumbers.beräkningar + 1 Then tempCol = ColumnNumbers.beräkningar + 1
                    Me.Cells(tempRow, tempCol).Value = debetBelopp - kreditBelopp
                    
                    ' Summera alla värden från kolumn K och till höger
                    currentSaldo = GetSaldoFromTillfällighetsytan(ColumnNumbers.beräkningar, kontoNummer)
                    
                    
                    ' Uppdatera saldot i ordinarie saldokolumn
                    Me.Cells(bokföringsradNummer, ColumnNumbers.saldo).Value = currentSaldo
                    
                    ' Lägg till ta-bort-rad-knapp om den inte redan finns
                    If Me.Cells(bokföringsradNummer, ColumnNumbers.diff).Value = "" Then
                        AddDeleteButton bokföringsradNummer
                    End If
                Else
                    MsgBox "Konto " & kontoNummer & " hittades inte i tillfällighetsytan.", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    Next cell
End Sub
Function columnLetter(columnNumber As Integer) As String
    Dim columnBokstäver As String
    columnBokstäver = ""
    
    Do While columnNumber > 0
        columnBokstäver = Chr((columnNumber - 1) Mod 26 + 65) & columnLetter
        columnNumber = (columnNumber - 1) \ 26
    Loop
    
    columnLetter = columnBokstäver
End Function
Private Sub UpdateDiff()
    Dim lastRow As Long
    Dim diffRow As Long
    Dim debetSum As Double
    Dim kreditSum As Double
    Dim debetColLetter As String
    Dim kreditColLetter As String
    
    ' Bestäm sista ifyllda raden i kolumn 1 (antagande att kolumn 1 alltid har data)
    lastRow = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row
    diffRow = lastRow + 10
    
    ' Hämta kolumnbokstäver för Debet och Kredit
    debetColLetter = columnLetter(ColumnNumbers.debet)
    kreditColLetter = columnLetter(ColumnNumbers.kredit)
    
    ' Använd kolumnbokstäverna i Range-funktionen för att summera debet och kredit
    debetSum = Application.WorksheetFunction.Sum(Me.Range(debetColLetter & "2:" & debetColLetter & lastRow))
    kreditSum = Application.WorksheetFunction.Sum(Me.Range(kreditColLetter & "2:" & kreditColLetter & lastRow))
    
    ' Ta bort tidigare diff-beräkningar från rad 9 under sista ifyllda raden
    Me.Rows(lastRow + 9).ClearContents
    
    ' Lägg till ny diff-beräkning på rad 10 under sista ifyllda raden
    Me.Cells(diffRow, ColumnNumbers.diff - 1).Value = "Diff"
    Me.Cells(diffRow, ColumnNumbers.diff).Value = debetSum - kreditSum
End Sub

' Hämta benämningen för ett konto
Private Function GetBenamning(ws As Worksheet, kontoNummer As String) As String
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow ' Förutsätter att den första raden är headers
        If ws.Cells(i, "G").Value = kontoNummer Then
            GetBenamning = ws.Cells(i, "H").Value
            Exit Function
        End If
    Next i
    GetBenamning = ""
End Function

' Hämta raden för ett konto
Private Function GetAccountRow(ws As Worksheet, kontoNummer As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow ' Förutsätter att den första raden är headers
        If ws.Cells(i, "G").Value = kontoNummer Then
            GetAccountRow = i
            Exit Function
        End If
    Next i
    GetAccountRow = 0
End Function

' Hämta raden för ett konto i tillfällighetsytan
Private Function GetAccountRowFromTillfällighetsytan(ws As Worksheet, kontoNummer As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, ColumnNumbers.kontoförändringar).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow ' Förutsätter att tillfällighetsytan börjar på rad 2
        If ws.Cells(i, ColumnNumbers.kontoförändringar).Value = kontoNummer Then
            GetAccountRowFromTillfällighetsytan = i
            Exit Function
        End If
    Next i
    GetAccountRowFromTillfällighetsytan = 0
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

Function GetSaldoFromTillfällighetsytan(startCol As Long, kontoNummer As String) As Double
    Dim tempRow As Long
    Dim tempCol As Long
    Dim saldo As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Bokföring")
    
    ' Hitta raden för kontonumret i tillfällighetsytan
    tempRow = GetAccountRowFromTillfällighetsytan(ws, kontoNummer)
    
    ' Kontrollera om raden hittades
    If tempRow > 0 Then
        saldo = ws.Cells(tempRow, startCol).Value
        tempCol = startCol + 1
        
        ' Summera alla värden från startCol och till höger
        Do While ws.Cells(tempRow, tempCol).Value <> ""
            saldo = saldo + ws.Cells(tempRow, tempCol).Value
            tempCol = tempCol + 1
        Loop
    Else
        saldo = 0 ' Om kontonumret inte hittades, sätt saldo till 0
    End If
    
    GetSaldoFromTillfällighetsytan = saldo
End Function

Function GetPreviousEndSaldo(kontoNummer As String, currentRow As Long) As Double
    Dim wsBokforing As Worksheet
    Dim saldo As Double
    Dim i As Long
    
    Set wsBokforing = ThisWorkbook.Sheets("Bokföring")
    
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

    Set ws = ThisWorkbook.Sheets("Bokföring")

    ' Definiera namn och titel för knappen
    btnName = "DeleteButton_" & rowNumber
    btnCaption = "Delete"

    ' Kontrollera om knappen redan finns på raden
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

        ' Ställ in knappens egenskaper
        With btn
            .Name = btnName
            .OnAction = "DeleteRow"
            .TextFrame.Characters.Text = btnCaption
        End With
    End If
End Sub


Public Sub UpdateTillfällighetsytanEfterBorttagAvRad(kontoNummer As String)
    Dim ws As Worksheet
    Dim tempRow As Long
    Dim lastTempCol As Long

    Set ws = ThisWorkbook.Sheets("Bokföring")
    
    ' Hitta raden för kontonumret i tillfällighetsytan
    tempRow = GetAccountRowFromTillfällighetsytan(ws, kontoNummer)
    Debug.Print "tempRow för konto " & kontoNummer & ": " & tempRow
    
    If tempRow > 0 Then
        ' Hitta sista kolumnen med ett värde i tillfällighetsytan
        lastTempCol = ws.Cells(tempRow, ws.Columns.Count).End(xlToLeft).Column
        Debug.Print "lastTempCol: " & lastTempCol
        
        ' Kolla om vi är i beräkningsområdet
        If lastTempCol >= ColumnNumbers.beräkningar Then
            ' Lägg till felsökning för att se vad som orsakar problemet
            Debug.Print "tempRow: " & tempRow
            Debug.Print "lastTempCol: " & lastTempCol
            Debug.Print "Cell Value: " & ws.Cells(tempRow, lastTempCol).Value
            
            ' Ta bort sista värdet i tillfällighetsytan
            ws.Cells(tempRow, lastTempCol).Value = ""
        End If
    Else
        ' Om raden inte hittas, skriv ut ett meddelande för felsökning
        Debug.Print "Kontonumret " & kontoNummer & " hittades inte i tillfällighetsytan."
    End If
End Sub


