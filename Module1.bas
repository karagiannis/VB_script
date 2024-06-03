Attribute VB_Name = "Module1"
    Sub Button1_Click()
    Debug.Print "Button click"
    BokforingKnapp_Click
End Sub

Sub BokforingKnapp_Click()
    If KontrolleraKrav() Then
        Dim lastRow As Long
        Dim i As Long
        
        ' Hämta sista ifyllda raden i Bokföringsbladet
        lastRow = Sheet5.Cells(Sheet5.Rows.Count, ColumnNumbers.Konto).End(xlUp).Row
        
        ' Uppdatera huvudboken för varje rad i Bokföringsbladet
        For i = 2 To lastRow
            Dim kontoNummer As String
            kontoNummer = Sheet5.Cells(i, ColumnNumbers.Konto).Value
            Debug.Print "Uppdaterar huvudbok för rad: " & i
            If kontoNummer <> "" Then
            Debug.Print "Kontonummer är:" & kontoNummer
                UppdateraHuvudbok kontoNummer, i
            End If
        Next i
        
        ' Uppdatera Verifikationslistan
        Debug.Print "Uppdaterar verifikationslista"
        UppdateraVerifikationslista
        
        ' Rensa Bokföringsbladet
        Debug.Print "Rensar bokföringsblad"
        RensaBokforingsblad
        
        MsgBox "Bokföring genomförd.", vbInformation
    End If
End Sub


Sub UppdateraHuvudbok2(kontoNummer As String, radNummer As Long)
    Dim wsAccount As Worksheet
    Set wsAccount = ThisWorkbook.Sheets(kontoNummer)
    Debug.Print "Arbetsblad satt till: " & wsAccount.Name
    
    ' Hämta senaste saldo
    Dim saldo As Double
    saldo = wsAccount.Cells(wsAccount.Cells(wsAccount.Rows.Count, 1).End(xlUp).Row, ColumnNumbers.saldo).Value
    Debug.Print "Hämtat saldo: " & saldo
    
    
    ' Hämta data från bokföringsbladet
    Dim verifikationsserie As String
    Dim verNr As String
    Dim systemdatum As String
    Dim registreringsdatum As String
    Dim kostnadsställe As String
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
    kostnadsställe = Sheet5.Cells(radNummer, ColumnNumbers.kostnadsställe).Value
    projekt = Sheet5.Cells(radNummer, ColumnNumbers.projekt).Value
    verifikationstext = Sheet5.Cells(radNummer, ColumnNumbers.verifikationstext).Value
    transaktionsinfo = Sheet5.Cells(radNummer, ColumnNumbers.transaktionsinfo).Value
    debet = Sheet5.Cells(radNummer, ColumnNumbers.debet).Value
    kredit = Sheet5.Cells(radNummer, ColumnNumbers.kredit).Value
    nyttSaldo = saldo + debet - kredit
    
    Debug.Print "Hämtat data: verifikationsserie=" & verifikationsserie & ", verNr=" & verNr & _
                ", systemdatum=" & systemdatum & ", registreringsdatum=" & registreringsdatum & _
                ", kostnadsställe=" & kostnadsställe & ", projekt=" & projekt & _
                ", verifikationstext=" & verifikationstext & ", transaktionsinfo=" & transaktionsinfo & _
                ", debet=" & debet & ", kredit=" & kredit & ", nyttSaldo=" & nyttSaldo
    
    ' Hitta nästa lediga rad i huvudboken
    Dim newRow As Long
    newRow = wsAccount.Cells(wsAccount.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Infoga data i huvudboken
    wsAccount.Cells(newRow, ColumnNumbers.Konto).Value = kontoNummer
    wsAccount.Cells(newRow, ColumnNumbers.Benämning).Value = Sheet5.Cells(radNummer, ColumnNumbers.Benämning).Value
    wsAccount.Cells(newRow, ColumnNumbers.verifikationsserie).Value = verifikationsserie
    wsAccount.Cells(newRow, ColumnNumbers.verNr).Value = verNr
    wsAccount.Cells(newRow, ColumnNumbers.systemdatum).Value = systemdatum
    wsAccount.Cells(newRow, ColumnNumbers.registreringsdatum).Value = registreringsdatum
    wsAccount.Cells(newRow, ColumnNumbers.kostnadsställe).Value = kostnadsställe
    wsAccount.Cells(newRow, ColumnNumbers.projekt).Value = projekt
    wsAccount.Cells(newRow, ColumnNumbers.verifikationstext).Value = verifikationstext
    wsAccount.Cells(newRow, ColumnNumbers.transaktionsinfo).Value = transaktionsinfo
    wsAccount.Cells(newRow, ColumnNumbers.debet).Value = debet
    wsAccount.Cells(newRow, ColumnNumbers.kredit).Value = kredit
    wsAccount.Cells(newRow, ColumnNumbers.saldo).Value = nyttSaldo
    wsAccount.Cells(newRow, ColumnNumbers.bokföringsunderlag).Value = Sheet5.Cells(radNummer, ColumnNumbers.bokföringsunderlag).Value
    
    Debug.Print "UppdateraHuvudbok avslutas för konto: " & kontoNummer & " och radnummer: " & radNummer
End Sub

Sub UppdateraHuvudbok(kontoNummer As String, radNummer As Long)
    Dim wsAccount As Worksheet
    Set wsAccount = ThisWorkbook.Sheets(kontoNummer)
    Debug.Print "Arbetsblad satt till: " & wsAccount.Name
    
    ' Hämta senaste saldo
    Dim saldo As Double
    saldo = wsAccount.Cells(wsAccount.Cells(wsAccount.Rows.Count, 1).End(xlUp).Row, ColumnNumbers.saldo).Value
    Debug.Print "Hämtat saldo: " & saldo
    
    ' Beräkna nytt saldo
    Dim nyttSaldo As Double
    nyttSaldo = saldo + Sheet5.Cells(radNummer, ColumnNumbers.debet).Value - Sheet5.Cells(radNummer, ColumnNumbers.kredit).Value
    
    ' Hitta nästa lediga rad i huvudboken
    Dim newRow As Long
    newRow = wsAccount.Cells(wsAccount.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Infoga data i huvudboken
    With wsAccount
        .Cells(newRow, ColumnNumbers.Konto).Value = kontoNummer
        .Cells(newRow, ColumnNumbers.Benämning).Value = Sheet5.Cells(radNummer, ColumnNumbers.Benämning).Value
        .Cells(newRow, ColumnNumbers.verifikationsserie).Value = Sheet5.Cells(radNummer, ColumnNumbers.verifikationsserie).Value
        .Cells(newRow, ColumnNumbers.verNr).Value = Sheet5.Cells(radNummer, ColumnNumbers.verNr).Value
        .Cells(newRow, ColumnNumbers.systemdatum).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
        .Cells(newRow, ColumnNumbers.registreringsdatum).Value = Sheet5.Cells(radNummer, ColumnNumbers.registreringsdatum).Value
        .Cells(newRow, ColumnNumbers.kostnadsställe).Value = Sheet5.Cells(radNummer, ColumnNumbers.kostnadsställe).Value
        .Cells(newRow, ColumnNumbers.projekt).Value = Sheet5.Cells(radNummer, ColumnNumbers.projekt).Value
        .Cells(newRow, ColumnNumbers.verifikationstext).Value = Sheet5.Cells(radNummer, ColumnNumbers.verifikationstext).Value
        .Cells(newRow, ColumnNumbers.transaktionsinfo).Value = Sheet5.Cells(radNummer, ColumnNumbers.transaktionsinfo).Value
        .Cells(newRow, ColumnNumbers.debet).Value = Sheet5.Cells(radNummer, ColumnNumbers.debet).Value
        .Cells(newRow, ColumnNumbers.kredit).Value = Sheet5.Cells(radNummer, ColumnNumbers.kredit).Value
        .Cells(newRow, ColumnNumbers.saldo).Value = nyttSaldo

        ' Använd funktionen för att kopiera hyperlänken
        KopieraHyperlänk Sheet5.Cells(radNummer, ColumnNumbers.bokföringsunderlag), .Cells(newRow, ColumnNumbers.bokföringsunderlag)
    End With
    
    Debug.Print "UppdateraHuvudbok avslutas för konto: " & kontoNummer & " och radnummer: " & radNummer
End Sub

Sub UppdateraVerifikationslista2()
    Dim verifikationsserie As String
    Dim verNr As String
    Dim systemdatum As String
    Dim registreringsdatum As String
    Dim kostnadsställe As String
    Dim projekt As String
    Dim verifikationstext As String
    Dim transaktionsinfo As String
    Dim debet As Double
    Dim kredit As Double
    Dim saldo As Double
    Dim diff As Double
    Dim bokföringsunderlag As String
    Dim kontoförändringar As String
    Dim beräkningar(1 To 6) As Double

    Dim wsVerifikationslista As Worksheet
    Set wsVerifikationslista = ThisWorkbook.Sheets("Verifikationslista")
    
    ' Hitta nästa lediga rad i Verifikationslistan
    Dim newRow As Long
    newRow = wsVerifikationslista.Cells(wsVerifikationslista.Rows.Count, 1).End(xlUp).Row + 1
    Debug.Print "Den funna lediga raden är i verifikationslistan är:" & newRow
    
    ' Hitta sista raden i Bokföring
    Dim lastRow As Long
    lastRow = Sheet5.Cells(Sheet5.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Den sista raden i Bokföring är:" & lastRow
    
    ' Samla data för varje rad
    Dim i As Long, j As Long
    For i = 2 To lastRow
        verifikationsserie = Sheet5.Cells(i, ColumnNumbers.verifikationsserie).Value
        verNr = Sheet5.Cells(i, ColumnNumbers.verNr).Value
        systemdatum = Format(Now, "yyyy-mm-dd hh:mm:ss")
        registreringsdatum = Sheet5.Cells(i, ColumnNumbers.registreringsdatum).Value
        kostnadsställe = Sheet5.Cells(i, ColumnNumbers.kostnadsställe).Value
        projekt = Sheet5.Cells(i, ColumnNumbers.projekt).Value
        verifikationstext = Sheet5.Cells(i, ColumnNumbers.verifikationstext).Value
        transaktionsinfo = Sheet5.Cells(i, ColumnNumbers.transaktionsinfo).Value
        debet = Sheet5.Cells(i, ColumnNumbers.debet).Value
        kredit = Sheet5.Cells(i, ColumnNumbers.kredit).Value
        saldo = Sheet5.Cells(i, ColumnNumbers.saldo).Value
        diff = Sheet5.Cells(i, ColumnNumbers.diff).Value
        bokföringsunderlag = Sheet5.Cells(i, ColumnNumbers.bokföringsunderlag).Value
        kontoförändringar = Sheet5.Cells(i, ColumnNumbers.kontoförändringar).Value
        ' Dim bokforingsunderlag As String
        ' MsgBox bokforingsunderlag

        
        Debug.Print "Hämtat data: verifikationsserie=" & verifikationsserie & ", verNr=" & verNr & _
                ", systemdatum=" & systemdatum & ", registreringsdatum=" & registreringsdatum & _
                ", kostnadsställe=" & kostnadsställe & ", projekt=" & projekt & _
                ", verifikationstext=" & verifikationstext & ", transaktionsinfo=" & transaktionsinfo & _
                ", debet=" & debet & ", kredit=" & kredit & ", saldo=" & saldo & _
                ", diff=" & diff & ", bokföringsunderlag =" & bokföringsunderlag & ", & kontoförändringar=" & kontoförändringar

    
        
        ' Hämta upp till 6 kolumner av beräkningar
        For j = 1 To 6
            beräkningar(j) = Sheet5.Cells(i, ColumnNumbers.beräkningar + j - 1).Value
             Debug.Print "Beräkning " & j & ": " & beräkningar(j)
        Next j
        
        ' Infoga data i Verifikationslistan
        With wsVerifikationslista
            .Cells(newRow, ColumnNumbers.Konto).Value = Sheet5.Cells(i, ColumnNumbers.Konto).Value
            .Cells(newRow, ColumnNumbers.Benämning).Value = Sheet5.Cells(i, ColumnNumbers.Benämning).Value
            .Cells(newRow, ColumnNumbers.Beskrivning).Value = Sheet5.Cells(i, ColumnNumbers.Beskrivning).Value
            .Cells(newRow, ColumnNumbers.verifikationsserie).Value = verifikationsserie
            .Cells(newRow, ColumnNumbers.verNr).Value = verNr
            .Cells(newRow, ColumnNumbers.systemdatum).Value = systemdatum
            .Cells(newRow, ColumnNumbers.registreringsdatum).Value = registreringsdatum
            .Cells(newRow, ColumnNumbers.kostnadsställe).Value = kostnadsställe
            .Cells(newRow, ColumnNumbers.projekt).Value = projekt
            .Cells(newRow, ColumnNumbers.verifikationstext).Value = verifikationstext
            .Cells(newRow, ColumnNumbers.transaktionsinfo).Value = transaktionsinfo
            .Cells(newRow, ColumnNumbers.debet).Value = debet
            .Cells(newRow, ColumnNumbers.kredit).Value = kredit
            .Cells(newRow, ColumnNumbers.saldo).Value = saldo
            .Cells(newRow, ColumnNumbers.diff).Value = diff
            .Cells(newRow, ColumnNumbers.bokföringsunderlag).Value = bokföringsunderlag
            .Cells(newRow, ColumnNumbers.kontoförändringar).Value = kontoförändringar
            
            ' Infoga beräkningar i Verifikationslistan
            For j = 1 To 6
                .Cells(newRow, ColumnNumbers.beräkningar + j - 1).Value = beräkningar(j)
            Next j
        End With
        
        newRow = newRow + 1
    Next i
End Sub
Sub UppdateraVerifikationslista()
    Dim wsVerifikationslista As Worksheet
    Set wsVerifikationslista = ThisWorkbook.Sheets("Verifikationslista")
    
    ' Hitta nästa lediga rad i Verifikationslistan
    Dim newRow As Long
    newRow = wsVerifikationslista.Cells(wsVerifikationslista.Rows.Count, 1).End(xlUp).Row + 1
    Debug.Print "Den funna lediga raden är i verifikationslistan är:" & newRow
    
    ' Hitta sista raden i Bokföring
    Dim lastRow As Long
    lastRow = Sheet5.Cells(Sheet5.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Den sista raden i Bokföring är:" & lastRow
    
    ' Samla data för varje rad
    Dim i As Long, j As Long
    For i = 2 To lastRow
        ' Infoga data i Verifikationslistan
        With wsVerifikationslista
            .Cells(newRow, ColumnNumbers.Konto).Value = Sheet5.Cells(i, ColumnNumbers.Konto).Value
            .Cells(newRow, ColumnNumbers.Benämning).Value = Sheet5.Cells(i, ColumnNumbers.Benämning).Value
            .Cells(newRow, ColumnNumbers.Beskrivning).Value = Sheet5.Cells(i, ColumnNumbers.Beskrivning).Value
            .Cells(newRow, ColumnNumbers.verifikationsserie).Value = Sheet5.Cells(i, ColumnNumbers.verifikationsserie).Value
            .Cells(newRow, ColumnNumbers.verNr).Value = Sheet5.Cells(i, ColumnNumbers.verNr).Value
            .Cells(newRow, ColumnNumbers.systemdatum).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
            .Cells(newRow, ColumnNumbers.registreringsdatum).Value = Sheet5.Cells(i, ColumnNumbers.registreringsdatum).Value
            .Cells(newRow, ColumnNumbers.kostnadsställe).Value = Sheet5.Cells(i, ColumnNumbers.kostnadsställe).Value
            .Cells(newRow, ColumnNumbers.projekt).Value = Sheet5.Cells(i, ColumnNumbers.projekt).Value
            .Cells(newRow, ColumnNumbers.verifikationstext).Value = Sheet5.Cells(i, ColumnNumbers.verifikationstext).Value
            .Cells(newRow, ColumnNumbers.transaktionsinfo).Value = Sheet5.Cells(i, ColumnNumbers.transaktionsinfo).Value
            .Cells(newRow, ColumnNumbers.debet).Value = Sheet5.Cells(i, ColumnNumbers.debet).Value
            .Cells(newRow, ColumnNumbers.kredit).Value = Sheet5.Cells(i, ColumnNumbers.kredit).Value
            .Cells(newRow, ColumnNumbers.saldo).Value = Sheet5.Cells(i, ColumnNumbers.saldo).Value
            .Cells(newRow, ColumnNumbers.diff).Value = Sheet5.Cells(i, ColumnNumbers.diff).Value

            ' Använd funktionen för att kopiera hyperlänken
            KopieraHyperlänk Sheet5.Cells(i, ColumnNumbers.bokföringsunderlag), .Cells(newRow, ColumnNumbers.bokföringsunderlag)
            
            .Cells(newRow, ColumnNumbers.kontoförändringar).Value = Sheet5.Cells(i, ColumnNumbers.kontoförändringar).Value
            
            ' Infoga beräkningar i Verifikationslistan
            For j = 1 To 6
                .Cells(newRow, ColumnNumbers.beräkningar + j - 1).Value = Sheet5.Cells(i, ColumnNumbers.beräkningar + j - 1).Value
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
    Debug.Print "Sista raden i bokföringsbladet innan rensning: " & lastRow
    
    ' Ta bort alla Form-knappar på bokföringsbladet
    Dim shp As Shape
    For Each shp In Sheet5.Shapes
        If shp.FormControlType = xlButtonControl Then
            If shp.TopLeftCell.Column = ColumnNumbers.diff Then
                Debug.Print "Tar bort knapp: " & shp.Name
                shp.Delete
             End If
        End If
    Next shp
    
    If lastRow >= 2 Then
        ' Rensa alla rader i bokföringsbladet
        Sheet5.Rows("2:" & lastRow).ClearContents
        Debug.Print "Rader från 2 till " & lastRow & " rensade"
        
        ' Rensa diff-beräkningar, som är 10 rader under den sista ifyllda raden
        For i = lastRow + 1 To lastRow + 10
            Sheet5.Rows(i).ClearContents
            Debug.Print "Rensade rad: " & i
        Next i
    End If
    
    InitializeBokföring
    Debug.Print "Bokföringsblad initierat"
    
    Debug.Print "RensaBokforingsblad avslutas"
End Sub



Function KontrolleraKrav() As Boolean
    KontrolleraKrav = True ' Anta att alla krav är uppfyllda
    
    ' Kontrollera att diff är 0
    Dim lastRow As Long
    lastRow = Sheet5.Cells(Sheet5.Rows.Count, ColumnNumbers.Konto).End(xlUp).Row
    Dim diffRow As Long
    diffRow = lastRow + 10
    
    If Sheet5.Cells(diffRow, ColumnNumbers.diff).Value <> 0 Then
        MsgBox "Diff måste vara 0 för att bokföringen ska kunna genomföras.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kontrollera att registreringsdatum är ifyllt
    If IsEmpty(Sheet5.Cells(2, ColumnNumbers.registreringsdatum).Value) Then
        MsgBox "Registreringsdatum måste vara ifyllt.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kontrollera att verifikationstext är ifyllt
    If IsEmpty(Sheet5.Cells(2, ColumnNumbers.verifikationstext).Value) Then
        MsgBox "Verifikationstext måste vara ifyllt.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kontrollera att minst två rader är ifyllda i bokföringsposten
    If lastRow < 3 Then
        MsgBox "Minst två rader måste vara ifyllda i bokföringsposten.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kontrollera att verifikationsserie och verifikationsnummer är ifyllda
    If IsEmpty(Sheet5.Cells(2, ColumnNumbers.verifikationsserie).Value) Or IsEmpty(Sheet5.Cells(2, ColumnNumbers.verNr).Value) Then
        MsgBox "Verifikationsserie och verifikationsnummer måste vara ifyllda.", vbExclamation
        KontrolleraKrav = False
    End If
    
    ' Kopiera gemensamma poster till varje rad från rad 2
    KopieraGemensammaPoster lastRow

End Function

Sub KopieraGemensammaPoster(lastRow As Long)
    Dim i As Long
    ' Loopa genom raderna från 3 till lastRow
    For i = 3 To lastRow
        ' Kopiera värden från rad 2 till nuvarande rad i
        Sheet5.Cells(i, ColumnNumbers.verifikationsserie).Value = Sheet5.Cells(2, ColumnNumbers.verifikationsserie).Value
        Sheet5.Cells(i, ColumnNumbers.verNr).Value = Sheet5.Cells(2, ColumnNumbers.verNr).Value
        Sheet5.Cells(i, ColumnNumbers.registreringsdatum).Value = Sheet5.Cells(2, ColumnNumbers.registreringsdatum).Value
        Sheet5.Cells(i, ColumnNumbers.kostnadsställe).Value = Sheet5.Cells(2, ColumnNumbers.kostnadsställe).Value
        Sheet5.Cells(i, ColumnNumbers.projekt).Value = Sheet5.Cells(2, ColumnNumbers.projekt).Value
        Sheet5.Cells(i, ColumnNumbers.verifikationstext).Value = Sheet5.Cells(2, ColumnNumbers.verifikationstext).Value
        
        ' Använd funktionen för att kopiera hyperlänken från rad 2 till nuvarande rad i
        Call KopieraHyperlänk(Sheet5.Cells(2, ColumnNumbers.bokföringsunderlag), Sheet5.Cells(i, ColumnNumbers.bokföringsunderlag))
    Next i
End Sub


Public Sub DeleteRow()
    Dim btnName As String
    Dim rowNumber As Long
    Dim kontoNummer As String

    On Error GoTo ErrorHandler

    ' Identifiera den rad där knappen finns
    btnName = Application.Caller
    rowNumber = ThisWorkbook.Sheets("Bokföring").Shapes(btnName).TopLeftCell.Row

    ' Hämta kontonumret från den raden
    kontoNummer = ThisWorkbook.Sheets("Bokföring").Cells(rowNumber, ColumnNumbers.Konto).Value

    ' Ta bort raden
    ThisWorkbook.Sheets("Bokföring").Rows(rowNumber).Delete

    ' Ta bort tillhörande knappar
    On Error Resume Next
    ThisWorkbook.Sheets("Bokföring").Shapes(btnName).Delete
    On Error GoTo 0

    ' Uppdatera tillfällighetsytan
    ThisWorkbook.Sheets("Bokföring").UpdateTillfällighetsytanEfterBorttagAvRad kontoNummer

    Exit Sub

ErrorHandler:
    MsgBox "Ett fel inträffade: " & Err.Description, vbExclamation
End Sub

Function KopieraHyperlänk2(källCell As Range, målCell As Range)
    On Error GoTo ErrorHandler ' Starta felhantering
    
    ' Kontrollera om källcellen har en hyperlänk
    If källCell.Hyperlinks.Count > 0 Then
        ' Lägg till hyperlänk till målcell
        målCell.Parent.Hyperlinks.Add Anchor:=målCell, _
                                      Address:=källCell.Hyperlinks(1).Address, _
                                      TextToDisplay:=källCell.Text
    Else
        ' Kopiera bara värdet om det inte finns någon hyperlänk
        målCell.Value = källCell.Value
    End If

    ' Avsluta funktionen normalt
    Exit Function

ErrorHandler:
    MsgBox "Ett fel inträffade vid kopiering av hyperlänk: " & Err.Description, vbExclamation
End Function

Function KopieraHyperlänk(källCell As Range, målCell As Range)
    If källCell.Hyperlinks.Count > 0 Then
        ' Använd en mellanlagring av värden för att säkerställa korrekt kopiering
        Dim länkAddress As String
        Dim länkText As String
        
        länkAddress = källCell.Hyperlinks(1).Address
        länkText = källCell.Text
        
        ' Lägg till hyperlänken till målcell
        målCell.Parent.Hyperlinks.Add Anchor:=målCell, _
                                      Address:=länkAddress, _
                                      TextToDisplay:=länkText
    Else
        målCell.Value = källCell.Value
    End If
End Function

