Option Explicit

Sub GeneraTurni2027()

    Dim ws As Worksheet
    Dim ultimaRiga As Long, r As Long, rInizio As Long
    Dim dataRiga As Date, settimane As Long, cicloRep As Long
    Dim giorno As Long, festivo As Boolean
    Dim turno1 As String, turno2 As String, turno3 As String
    Dim rep(0 To 3) As String

    Set ws = ActiveSheet

    ' Cerca il 04/01/2027
    For r = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If IsDate(ws.Cells(r, "A").Value) Then
            If DateValue(ws.Cells(r, "A").Value) = DateSerial(2027, 1, 4) Then
                rInizio = r
                Exit For
            End If
        End If
    Next r

    If rInizio = 0 Then
        MsgBox "Non trovo il 04/01/2027 nella colonna A.", vbExclamation
        Exit Sub
    End If

    ' I nomi vengono letti dalla settimana iniziale
    turno1 = Trim(CStr(ws.Cells(rInizio, "C").Value))
    turno2 = Trim(CStr(ws.Cells(rInizio, "D").Value))
    turno3 = Trim(CStr(ws.Cells(rInizio, "E").Value))

    If turno1 = "" Or turno2 = "" Or turno3 = "" Then
        MsgBox "Compila C/D/E del 04/01/2027 prima di avviare la macro.", vbExclamation
        Exit Sub
    End If

    ' Reperibilita': 02/01 GALLININI, poi VENTURA, RUSSO, TANSINI
    rep(0) = "GALLININI"
    rep(1) = "VENTURA"
    rep(2) = "RUSSO"
    rep(3) = "TANSINI"

    ultimaRiga = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ==========================================================
    ' TURNI PRINCIPALI
    ' ==========================================================
    For r = rInizio To ultimaRiga

        If IsDate(ws.Cells(r, "A").Value) Then
            dataRiga = DateValue(ws.Cells(r, "A").Value)

            If dataRiga >= DateSerial(2027, 1, 4) Then

                festivo = RigaFestiva(ws, r)
                giorno = Weekday(dataRiga, vbMonday)

                If Not festivo And giorno <> 7 Then

                    settimane = DateDiff("d", DateSerial(2027, 1, 4), dataRiga) \ 7

                    Select Case settimane Mod 3

                        Case 0
                            ws.Cells(r, "C").Value = turno1
                            ws.Cells(r, "D").Value = turno2
                            ws.Cells(r, "E").Value = turno3

                        Case 1
                            ws.Cells(r, "C").Value = turno2
                            ws.Cells(r, "D").Value = turno3
                            ws.Cells(r, "E").Value = turno1

                        Case 2
                            ws.Cells(r, "C").Value = turno3
                            ws.Cells(r, "D").Value = turno1
                            ws.Cells(r, "E").Value = turno2

                    End Select

                    ApplicaRiposi ws, r, giorno, settimane Mod 3

                End If
            End If
        End If

    Next r

    ' ==========================================================
    ' REPERIBILITA' WEEKEND
    ' 02-03/01 GALLININI
    ' 09-10/01 VENTURA
    ' 16-17/01 RUSSO
    ' 23-24/01 TANSINI
    ' poi ripete
    ' ==========================================================
    For r = 1 To ultimaRiga

        If IsDate(ws.Cells(r, "A").Value) Then
            dataRiga = DateValue(ws.Cells(r, "A").Value)

            If dataRiga >= DateSerial(2027, 1, 2) _
               And dataRiga <= DateSerial(2027, 12, 31) Then

                giorno = Weekday(dataRiga, vbMonday)

                If giorno = 6 Or giorno = 7 Then

                    If Not RigaFestiva(ws, r) Then
                        cicloRep = (DateDiff("d", DateSerial(2027, 1, 2), dataRiga) \ 7) Mod 4
                        ws.Cells(r, "G").Value = rep(cicloRep)
                    End If

                End If
            End If
        End If

    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Turni, riposi e reperibilita' 2027 generati.", vbInformation

End Sub


Private Sub ApplicaRiposi(ByVal ws As Worksheet, ByVal r As Long, _
                          ByVal giorno As Long, ByVal ciclo As Long)

    Select Case ciclo

        ' SETTIMANA 1
        ' 1° TANSINI / 2° VENTURA / 3° RUSSO
        ' Mer: VENTURA riposa -> SAUNA sul 2°
        ' Gio: RUSSO riposa   -> SAUNA sul 3°
        ' Ven: TANSINI riposa -> SAUNA sul 1°
        Case 0
            Select Case giorno
                Case 3
                    ws.Cells(r, "D").Value = "SAUNA"
                Case 4
                    ws.Cells(r, "E").Value = "SAUNA"
                Case 5
                    ws.Cells(r, "C").Value = "SAUNA"
            End Select

        ' SETTIMANA 2
        ' 1° VENTURA / 2° RUSSO / 3° TANSINI
        ' Gio: VENTURA riposa; RUSSO copre il 1°; SAUNA il 2°
        ' Ven: RUSSO e TANSINI riposano;
        '      GALLININI copre il 1°; SAUNA copre il 2°.
        Case 1
            Select Case giorno
                Case 4
                    ws.Cells(r, "C").Value = "RUSSO"
                    ws.Cells(r, "D").Value = "SAUNA"
                Case 5
                    ws.Cells(r, "C").Value = "GALLININI"
                    ws.Cells(r, "D").Value = "SAUNA"
            End Select

        ' SETTIMANA 3
        ' 1° RUSSO / 2° TANSINI / 3° VENTURA
        ' Gio: TANSINI riposa -> SAUNA sul 2°
        ' Ven: VENTURA riposa -> SAUNA sul 3°
        ' RUSSO non riposa.
        Case 2
            Select Case giorno
                Case 4
                    ws.Cells(r, "D").Value = "SAUNA"
                Case 5
                    ws.Cells(r, "E").Value = "SAUNA"
            End Select

    End Select

End Sub


Private Function RigaFestiva(ByVal ws As Worksheet, ByVal r As Long) As Boolean

    Dim testo As String

    testo = LCase(CStr(ws.Cells(r, "B").Value) & " " & _
                  CStr(ws.Cells(r, "C").Value) & " " & _
                  CStr(ws.Cells(r, "D").Value) & " " & _
                  CStr(ws.Cells(r, "E").Value))

    RigaFestiva = (InStr(1, testo, "festivit", vbTextCompare) > 0)

End Function
