Attribute VB_Name = "Module1"
Sub Button1_Click()
    Debug.Print "Button click"
    BokforingKnapp_Click
End Sub

Sub BokforingKnapp_Click()
    If KontrolleraKrav() Then
        Dim lastRow As Long
        Dim i As Long
        
        ' H�mta sista ifyllda raden i Bokf�ringsbladet
        lastRow = Sheet5.Cells(Sheet5.Rows.Count, ColumnNumbers.Konto).End(xlUp).Row
        
        ' Uppdatera huvudboken f�r varje rad i Bokf�ringsbladet
        For i = 2 To lastRow
            Dim kontoNummer As String
            kontoNummer = Sheet5.Cells(i, ColumnNumbers.Konto).Value
            Debug.Print "Uppdaterar huvudbok f�r rad: " & i
            If kontoNummer <> "" Then
            Debug.Print "Kontonummer �r:" & kontoNummer
                UppdateraHuvudbok kontoNummer, i
            End If
        Next i
        
        ' Uppdatera Verifikationslistan
        Debug.Print "Uppdaterar verifikationslista"
        UppdateraVerifikationslista
        
        ' Rensa Bokf�ringsbladet
        Debug.Print "Rensar bokf�ringsblad"
        RensaBokforingsblad
        
        MsgBox "Bokf�ring genomf�rd.", vbInformation
    End If
End Sub


Sub UppdateraHuvudbok(kontoNummer As String, radNummer As Long)
    Dim wsAccount As Worksheet
    Set wsAccount = ThisWorkbook.Sheets(kontoNummer)
    Debug.Print "Arbetsblad satt till: " & wsAccount.Name
    
    ' H�mta senaste saldo
    Dim saldo As Double
    saldo = wsAccount.Cells(wsAccount.Cells(wsAccount.Rows.Count, 1).End(xlUp).Row, ColumnNumbers.saldo).Value
    Debug.Print "H�mtat saldo: " & saldo
    
    
    ' H�mta data fr�n bokf�ringsbladet
    Dim verifikationsserie As String
    Dim verNr As String
    Dim systemdatum As String
    Dim registreringsdatum As String
    Dim kostnadsst�lle As String
    Dim projekt As String
    Dim verifikationstext As String
    Dim transaktionsinfo As String
    Dim debet As Double
    Dim kredit As Double
    Dim nyttSaldo As Double
    
    verifikationsserie = Sheet5.Cells(radNummer, ColumnNumbers.verifikationsserie).Value
    verNr = Sheet5.Cells(radNummer, ColumnNumbers.verNr).Value
    systemdatum = Format(Now, "yyyy-mm-dd hh:mm:ss")
    registreringsdatum = Sheet5.Cells(radNummer, ColumnNumbers.registreringsdatum).Value
    kostnadsst�lle = Sheet5.Cells(radNummer, ColumnNumbers.kostnadsst�lle).Value
    projekt = Sheet5.Cells(radNummer, ColumnNumbers.projekt).Value
    verifikationstext = Sheet5.Cells(radNummer, ColumnNumbers.verifikationstext).Value
    transaktionsinfo = Sheet5.Cells(radNummer, ColumnNumbers.transaktionsinfo).Value
    debet = Sheet5.Cells(radNummer, ColumnNumbers.debet).Value
    kredit = Sheet5.Cells(radNummer, ColumnNumbers.kredit).Value
    nyttSaldo = saldo + debet - kredit
    
    Debug.Print "H�mtat data: verifikationsserie=" & verifikationsserie & ", verNr=" & verNr & _
                ", systemdatum=" & systemdatum & ", registreringsdatum=" & registreringsdatum & _
                ", kostnadsst�lle=" & kostnadsst�lle & ", projekt=" & projekt & _
                ", verifikationstext=" & verifikationstext & ", transaktionsinfo=" & transaktionsinfo & _
                ", debet=" & debet & ", kredit=" & kredit & ", nyttSaldo=" & nyttSaldo
    
    ' Hitta n�sta lediga rad i huvudboken
    Dim newRow As Long
    newRow = wsAccount.Cells(wsAccount.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Infoga data i huvudboken
    wsAccount.Cells(newRow, ColumnNumbers.Konto).Value = kontoNummer
    wsAccount.Cells(newRow, ColumnNumbers.Ben�mning).Value = Sheet5.Cells(radNummer, ColumnNumbers.Ben�mning).Value
    wsAccount.Cells(newRow, ColumnNumbers.verifikationsserie).Value = verifikationsserie
    wsAccount.Cells(newRow, ColumnNumbers.verNr).Value = verNr
    wsAccount.Cells(newRow, ColumnNumbers.systemdatum).Value = systemdatum
    wsAccount.Cells(newRow, ColumnNumbers.registreringsdatum).Value = registreringsdatum
    wsAccount.Cells(newRow, ColumnNumbers.kostnadsst�lle).Value = kostnadsst�lle
    wsAccount.Cells(newRow, ColumnNumbers.projekt).Value = projekt
    wsAccount.Cells(newRow, ColumnNumbers.verifikationstext).Value = verifikationstext
    wsAccount.Cells(newRow, ColumnNumbers.transaktionsinfo).Value = transaktionsinfo
    wsAccount.Cells(newRow, ColumnNumbers.debet).Value = debet
    wsAccount.Cells(newRow, ColumnNumbers.kredit).Value = kredit
    wsAccount.Cells(newRow, ColumnNumbers.saldo).Value = nyttSaldo
    wsAccount.Cells(newRow, ColumnNumbers.bokf�ringsunderlag).Value = Sheet5.Cells(radNummer, ColumnNumbers.bokf�ringsunderlag).Value
    
    Debug.Print "UppdateraHuvudbok avslutas f�r konto: " & kontoNummer & " och radnummer: " & radNummer
End Sub
Sub UppdateraVerifikationslista()
    Dim verifikationsserie As String
    Dim verNr As String
    Dim systemdatum As String
    Dim registreringsdatum As String
    Dim kostnadsst�lle As String
    Dim projekt As String
    Dim verifikationstext As String
    Dim transaktionsinfo As String
    Dim debet As Double
    Dim kredit As Double
    Dim saldo As Double
    Dim diff As Double
    Dim bokf�ringsunderlag As String
    Dim kontof�r�ndringar As String
    Dim ber�kningar(1 To 6) As Double

    Dim wsVerifikationslista As Worksheet
    Set wsVerifikationslista = ThisWorkbook.Sheets("Verifikationslista")
    
    ' Hitta n�sta lediga rad i Verifikationslistan
    Dim newRow As Long
    newRow = wsVerifikationslista.Cells(wsVerifikationslista.Rows.Count, 1).End(xlUp).Row + 1
    Debug.Print "Den funna lediga raden �r i verifikationslistan �r:" & newRow
    
    ' Hitta sista raden i Bokf�ring
    Dim lastRow As Long
    lastRow = Sheet5.Cells(Sheet5.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Den sista raden i Bokf�ring �r:" & lastRow
    
    ' Samla data f�r varje rad
    Dim i As Long, j As Long
    For i = 2 To lastRow
        verifikationsserie = Sheet5.Cells(i, ColumnNumbers.verifikationsserie).Value
        verNr = Sheet5.Cells(i, ColumnNumbers.verNr).Value
        systemdatum = Format(Now, "yyyy-mm-dd hh:mm:ss")
        registreringsdatum = Sheet5.Cells(i, ColumnNumbers.registreringsdatum).Value
        kostnadsst�lle = Sheet5.Cells(i, ColumnNumbers.kostnadsst�lle).Value
        projekt = Sheet5.Cells(i, ColumnNumbers.projekt).Value
        verifikationstext = Sheet5.Cells(i, ColumnNumbers.verifikationstext).Value
        transaktionsinfo = Sheet5.Cells(i, ColumnNumbers.transaktionsinfo).Value
        debet = Sheet5.Cells(i, ColumnNumbers.debet).Value
        kredit = Sheet5.Cells(i, ColumnNumbers.kredit).Value
        saldo = Sheet5.Cells(i, ColumnNumbers.saldo).Value
        diff = Sheet5.Cells(i, ColumnNumbers.diff).Value
        bokf�ringsunderlag = Sheet5.Cells(i, ColumnNumbers.bokf�ringsunderlag).Value
        kontof�r�ndringar = Sheet5.Cells(i, ColumnNumbers.kontof�r�ndringar).Value
        
        Debug.Print "H�mtat data: verifikationsserie=" & verifikationsserie & ", verNr=" & verNr & _
                ", systemdatum=" & systemdatum & ", registreringsdatum=" & registreringsdatum & _
                ", kostnadsst�lle=" & kostnadsst�lle & ", projekt=" & projekt & _
                ", verifikationstext=" & verifikationstext & ", transaktionsinfo=" & transaktionsinfo & _
                ", debet=" & debet & ", kredit=" & kredit & ", saldo=" & saldo & _
                ", diff=" & diff & ", bokf�ringsunderlag =" & bokf�ringsunderlag & ", & kontof�r�ndringar=" & kontof�r�ndringar

    
        
        ' H�mta upp till 6 kolumner av ber�kningar
        For j = 1 To 6
            ber�kningar(j) = Sheet5.Cells(i, ColumnNumbers.ber�kningar + j - 1).Value
             Debug.Print "Ber�kning " & j & ": " & ber�kningar(j)
        Next j
        
        ' Infoga data i Verifikationslistan
        With wsVerifikationslista
            .Cells(newRow, ColumnNumbers.Konto).Value = Sheet5.Cells(i, ColumnNumbers.Konto).Value
            .Cells(newRow, ColumnNumbers.Ben�mning).Value = Sheet5.Cells(i, ColumnNumbers.Ben�mning).Value
            .Cells(newRow, ColumnNumbers.Beskrivning).Value = Sheet5.Cells(i, ColumnNumbers.Beskrivning).Value
            .Cells(newRow, ColumnNumbers.verifikationsserie).Value = verifikationsserie
            .Cells(newRow, ColumnNumbers.verNr).Value = verNr
            .Cells(newRow, ColumnNumbers.systemdatum).Value = systemdatum
            .Cells(newRow, ColumnNumbers.registreringsdatum).Value = registreringsdatum
            .Cells(newRow, ColumnNumbers.kostnadsst�lle).Value = kostnadsst�lle
            .Cells(newRow, ColumnNumbers.projekt).Value = projekt
            .Cells(newRow, ColumnNumbers.verifikationstext).Value = verifikationstext
            .Cells(newRow, ColumnNumbers.transaktionsinfo).Value = transaktionsinfo
            .Cells(newRow, ColumnNumbers.debet).Value = debet
            .Cells(newRow, ColumnNumbers.kredit).Value = kredit
            .Cells(newRow, ColumnNumbers.saldo).Value = saldo
            .Cells(newRow, ColumnNumbers.diff).Value = diff
            .Cells(newRow, ColumnNumbers.bokf�ringsunderlag).Value = bokf�ringsunderlag
            .Cells(newRow, ColumnNumbers.kontof�r�ndringar).Value = kontof�r�ndringar
            
            ' Infoga ber�kningar i Verifikationslistan
            For j = 1 To 6
                .Cells(newRow, ColumnNumbers.ber�kningar + j - 1).Value = ber�kningar(j)
            Next j
        End With
        
        newRow = newRow + 1
    Next i
End Sub
Sub RensaBokforingsblad()
    Debug.Print "RensaBokforingsblad startar"
    
    Dim lastRow As Long
    Dim i As Long
    lastRow = Sheet5.Cells(Sheet5.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Sista raden i bokf�ringsbladet innan rensning: " & lastRow
    
    ' Ta bort alla Form-knappar p� bokf�ringsbladet
    Dim shp As Shape
    For Each shp In Sheet5.Shapes
        If shp.FormControlType = xlButtonControl Then
            Debug.Print "Tar bort knapp: " & shp.Name
            shp.Delete
        End If
    Next shp
    
    If lastRow >= 2 Then
        ' Rensa alla rader i bokf�ringsbladet
        Sheet5.Rows("2:" & lastRow).ClearContents
        Debug.Print "Rader fr�n 2 till " & lastRow & " rensade"
        
        ' Rensa diff-ber�kningar, som �r 10 rader under den sista ifyllda raden
        For i = lastRow + 1 To lastRow + 10
            Sheet5.Rows(i).ClearContents
            Debug.Print "Rensade rad: " & i
        Next i
    End If
    
    InitializeBokf�ring
    Debug.Print "Bokf�ringsblad initierat"
    
    Debug.Print "RensaBokforingsblad avslutas"
End Sub



Function KontrolleraKrav() As Boolean
    KontrolleraKrav = True ' Anta att alla krav �r uppfyllda
    
    ' Kontrollera att diff �r 0
    Dim lastRow As Long
    lastRow = Sheet5.Cells(Sheet5.Rows.Count, ColumnNumbers.Konto).End(xlUp).Row
    Dim diffRow As Long
    diffRow = lastRow + 10
    
    If Sheet5.Cells(diffRow, ColumnNumbers.diff).Value <> 0 Then
        MsgBox "Diff m�ste vara 0 f�r att bokf�ringen ska kunna genomf�ras.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kontrollera att registreringsdatum �r ifyllt
    If IsEmpty(Sheet5.Cells(2, ColumnNumbers.registreringsdatum).Value) Then
        MsgBox "Registreringsdatum m�ste vara ifyllt.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kontrollera att verifikationstext �r ifyllt
    If IsEmpty(Sheet5.Cells(2, ColumnNumbers.verifikationstext).Value) Then
        MsgBox "Verifikationstext m�ste vara ifyllt.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kontrollera att minst tv� rader �r ifyllda i bokf�ringsposten
    If lastRow < 3 Then
        MsgBox "Minst tv� rader m�ste vara ifyllda i bokf�ringsposten.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kontrollera att verifikationsserie och verifikationsnummer �r ifyllda
    If IsEmpty(Sheet5.Cells(2, ColumnNumbers.verifikationsserie).Value) Or IsEmpty(Sheet5.Cells(2, ColumnNumbers.verNr).Value) Then
        MsgBox "Verifikationsserie och verifikationsnummer m�ste vara ifyllda.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kopiera gemensamma poster till varje rad fr�n rad 2
    KopieraGemensammaPoster lastRow

End Function


Sub KopieraGemensammaPoster(lastRow As Long)
    Dim i As Long
    For i = 3 To lastRow
        Sheet5.Cells(i, ColumnNumbers.verifikationsserie).Value = Sheet5.Cells(2, ColumnNumbers.verifikationsserie).Value
        Sheet5.Cells(i, ColumnNumbers.verNr).Value = Sheet5.Cells(2, ColumnNumbers.verNr).Value
        Sheet5.Cells(i, ColumnNumbers.registreringsdatum).Value = Sheet5.Cells(2, ColumnNumbers.registreringsdatum).Value
        Sheet5.Cells(i, ColumnNumbers.kostnadsst�lle).Value = Sheet5.Cells(2, ColumnNumbers.kostnadsst�lle).Value
        Sheet5.Cells(i, ColumnNumbers.projekt).Value = Sheet5.Cells(2, ColumnNumbers.projekt).Value
        Sheet5.Cells(i, ColumnNumbers.verifikationstext).Value = Sheet5.Cells(2, ColumnNumbers.verifikationstext).Value
    Next i
End Sub



Public Sub DeleteRow()
    Dim btnName As String
    Dim rowNumber As Long
    Dim kontoNummer As String

    On Error GoTo ErrorHandler

    ' Identifiera den rad d�r knappen finns
    btnName = Application.Caller
    rowNumber = ThisWorkbook.Sheets("Bokf�ring").Shapes(btnName).TopLeftCell.Row

    ' H�mta kontonumret fr�n den raden
    kontoNummer = ThisWorkbook.Sheets("Bokf�ring").Cells(rowNumber, ColumnNumbers.Konto).Value

    ' Ta bort raden
    ThisWorkbook.Sheets("Bokf�ring").Rows(rowNumber).Delete

    ' Ta bort tillh�rande knappar
    On Error Resume Next
    ThisWorkbook.Sheets("Bokf�ring").Shapes(btnName).Delete
    On Error GoTo 0

    ' Uppdatera tillf�llighetsytan
    ThisWorkbook.Sheets("Bokf�ring").UpdateTillf�llighetsytanEfterBorttagAvRad kontoNummer

    Exit Sub

ErrorHandler:
    MsgBox "Ett fel intr�ffade: " & Err.Description, vbExclamation
End Sub


