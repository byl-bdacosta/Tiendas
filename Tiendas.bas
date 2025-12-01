Option Explicit

'====================================================
'   DEFINICIÓN DEL TIPO DE HORARIO POR DÍA
'====================================================
Type HorarioDia
    Inicio As Variant
    PartidoInicio As Variant
    PartidoFin As Variant
    Fin As Variant
End Type

'====================================================
'   FORMATEADOR DE HORAS (EVITA 0,41666…)
'====================================================
Private Function FmtHora(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or v = "" Then
        FmtHora = ""
    Else
        FmtHora = Format$(v, "hh:mm")
    End If
End Function

'====================================================
'   COMPARAR DOS DÍAS COMPLETOS
'====================================================
Private Function HorarioIgual(h1 As HorarioDia, h2 As HorarioDia) As Boolean
    HorarioIgual = (h1.Inicio = h2.Inicio And _
                    h1.Fin = h2.Fin And _
                    h1.PartidoInicio = h2.PartidoInicio And _
                    h1.PartidoFin = h2.PartidoFin)
End Function

'====================================================
'   SABER SI UN DÍA ESTÁ VACÍO (CERRADO)
'====================================================
Private Function DiaVacio(h As HorarioDia) As Boolean
    DiaVacio = (FmtHora(h.Inicio) = "" And _
                FmtHora(h.Fin) = "" And _
                FmtHora(h.PartidoInicio) = "" And _
                FmtHora(h.PartidoFin) = "")
End Function

'====================================================
'   DETERMINAR EL CASO (1-5)
'====================================================
Private Function ObtenerCaso(L As HorarioDia, S As HorarioDia, D As HorarioDia) As Integer

    'CASO 1 – MISMO HORARIO TODOS LOS DÍAS
    If HorarioIgual(L, S) And HorarioIgual(L, D) Then
        ObtenerCaso = 1: Exit Function
    End If

    'CASO 2 – MISMO Lun-Sáb, Domingo distinto o cerrado
    If HorarioIgual(L, S) And Not HorarioIgual(L, D) Then
        ObtenerCaso = 2: Exit Function
    End If

    'CASO 3 – Jornada partida Lun-Vie y Sábado, Domingo cerrado
    If FmtHora(L.PartidoInicio) <> "" And FmtHora(S.PartidoInicio) <> "" And DiaVacio(D) Then
        ObtenerCaso = 3: Exit Function
    End If

    'CASO 4 – Lun-Sáb continuo, Domingo continuo distinto
    If FmtHora(L.PartidoInicio) = "" And FmtHora(S.PartidoInicio) = "" And _
       FmtHora(D.PartidoInicio) = "" And HorarioIgual(L, S) And Not HorarioIgual(L, D) Then
        ObtenerCaso = 4: Exit Function
    End If

    'CASO 5 – Los tres días con partida
    If FmtHora(L.PartidoInicio) <> "" And FmtHora(S.PartidoInicio) <> "" And FmtHora(D.PartidoInicio) <> "" Then
        ObtenerCaso = 5: Exit Function
    End If

    'SI NO ENCAJA → CASO 0
    ObtenerCaso = 0
End Function

'====================================================
'   FORMATO DEL TEXTO (CASOS 1-5)
'====================================================
Private Function TextoHorario(caso As Integer, _
                              L As HorarioDia, S As HorarioDia, D As HorarioDia, _
                              idioma As String) As String
                              
    Dim sMonFri As String, sMonSat As String, sMonSun As String
    Dim sSat As String, sSun As String
    Dim sep As String: sep = " | "

    '----- LITERALES POR IDIOMA -----
    Select Case UCase(idioma)
        Case "EN"
            sMonFri = "Mon - Fri ": sMonSat = "Mon - Sat ": sMonSun = "Mon - Sun "
            sSat = "Sat: ": sSun = "Sun: "
        Case "ES"
            sMonFri = "Lun - Vie ": sMonSat = "Lun - Sáb ": sMonSun = "Lun - Dom "
            sSat = "Sáb: ": sSun = "Dom: "
        Case "GL"
            sMonFri = "Luns - Ven ": sMonSat = "Luns - Sáb ": sMonSun = "Luns - Dom "
            sSat = "Sáb: ": sSun = "Dom: "
        Case "CA"
            sMonFri = "Dl. - Dv. ": sMonSat = "Dl. - Ds. ": sMonSun = "Dl. - Dg. "
            sSat = "Ds.: ": sSun = "Dg.: "
        Case Else
            sMonFri = "Mon - Fri ": sMonSat = "Mon - Sat ": sMonSun = "Mon - Sun "
            sSat = "Sat: ": sSun = "Sun: "
    End Select

    '----- CASOS -----
    Select Case caso

        Case 1
            TextoHorario = sMonSun & FmtHora(L.Inicio) & " - " & FmtHora(L.Fin)

        Case 2
            If DiaVacio(D) Then
                TextoHorario = sMonSat & FmtHora(L.Inicio) & " - " & FmtHora(L.Fin)
            Else
                TextoHorario = sMonSat & FmtHora(L.Inicio) & " - " & FmtHora(L.Fin) & _
                               sep & sSun & FmtHora(D.Inicio) & " - " & FmtHora(D.Fin)
            End If

        Case 3
            TextoHorario = sMonFri & FmtHora(L.Inicio) & " - " & FmtHora(L.PartidoFin) & _
                           " / " & FmtHora(L.PartidoInicio) & " - " & FmtHora(L.Fin) & _
                           sep & sSat & FmtHora(S.Inicio) & " - " & FmtHora(S.PartidoFin) & _
                           " / " & FmtHora(S.PartidoInicio) & " - " & FmtHora(S.Fin)

        Case 4
            TextoHorario = sMonSat & FmtHora(L.Inicio) & " - " & FmtHora(L.Fin) & _
                           sep & sSun & FmtHora(D.Inicio) & " - " & FmtHora(D.Fin)

        Case 5
            TextoHorario = sMonFri & FmtHora(L.Inicio) & " - " & FmtHora(L.PartidoFin) & _
                           " / " & FmtHora(L.PartidoInicio) & " - " & FmtHora(L.Fin) & _
                           sep & sSat & FmtHora(S.Inicio) & " - " & FmtHora(S.PartidoFin) & _
                           " / " & FmtHora(S.PartidoInicio) & " - " & FmtHora(S.Fin) & _
                           sep & sSun & FmtHora(D.Inicio) & " - " & FmtHora(D.PartidoFin) & _
                           " / " & FmtHora(D.PartidoInicio) & " - " & FmtHora(D.Fin)

    End Select
End Function

'====================================================
'   FORMATO GENÉRICO SI NO HAY CASO
'====================================================
Private Function TextoHorarioGenerico(L As HorarioDia, S As HorarioDia, D As HorarioDia, idioma As String) As String
    Dim sMonFri As String, sSat As String, sSun As String
    Dim sep As String: sep = " | "

    Select Case UCase(idioma)
        Case "EN"
            sMonFri = "Mon - Fri ": sSat = "Sat: ": sSun = "Sun: "
        Case "ES"
            sMonFri = "Lun - Vie ": sSat = "Sáb: ": sSun = "Dom: "
        Case "GL"
            sMonFri = "Luns - Ven ": sSat = "Sáb: ": sSun = "Dom: "
        Case "CA"
            sMonFri = "Dl. - Dv. ": sSat = "Ds.: ": sSun = "Dg.: "
        Case Else
            sMonFri = "Mon - Fri ": sSat = "Sat: ": sSun = "Sun: "
    End Select

    TextoHorarioGenerico = _
        sMonFri & FmtHora(L.Inicio) & " - " & FmtHora(L.Fin) & _
        sep & sSat & FmtHora(S.Inicio) & " - " & FmtHora(S.Fin) & _
        sep & sSun & FmtHora(D.Inicio) & " - " & FmtHora(D.Fin)
End Function

'====================================================
'   MACRO PRINCIPAL
'====================================================
Public Sub Horarios()

    Dim ws As Worksheet
    Dim uFila As Long, i As Long
    Dim L As HorarioDia, S As HorarioDia, D As HorarioDia
    Dim caso As Integer

    Set ws = Sheets("Horarios habituales")

    uFila = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For i = 5 To uFila

        '------ Lunes-Viernes ------
        L.Inicio = ws.Cells(i, "C").Value
        If ws.Cells(i, "D").Value <> ws.Cells(i, "E").Value Then
            L.PartidoInicio = ws.Cells(i, "D").Value
            L.PartidoFin = ws.Cells(i, "E").Value
        Else
            L.PartidoInicio = Empty
            L.PartidoFin = Empty
        End If
        L.Fin = ws.Cells(i, "F").Value

        '------ Sábado ------
        S.Inicio = ws.Cells(i, "G").Value
        If ws.Cells(i, "H").Value <> ws.Cells(i, "I").Value Then
            S.PartidoInicio = ws.Cells(i, "H").Value
            S.PartidoFin = ws.Cells(i, "I").Value
        Else
            S.PartidoInicio = Empty
            S.PartidoFin = Empty
        End If
        S.Fin = ws.Cells(i, "J").Value

        '------ Domingo ------
        D.Inicio = ws.Cells(i, "K").Value
        If ws.Cells(i, "L").Value <> ws.Cells(i, "M").Value Then
            D.PartidoInicio = ws.Cells(i, "L").Value
            D.PartidoFin = ws.Cells(i, "M").Value
        Else
            D.PartidoInicio = Empty
            D.PartidoFin = Empty
        End If
        D.Fin = ws.Cells(i, "N").Value

        '------ Determinar caso ------
        caso = ObtenerCaso(L, S, D)

        '------ Escribir en Inglés / Español / Gallego / Catalán ------
        If caso = 0 Then
            ws.Cells(i, "AF").Value = TextoHorarioGenerico(L, S, D, "EN")
            ws.Cells(i, "AG").Value = TextoHorarioGenerico(L, S, D, "ES")
            ws.Cells(i, "AH").Value = TextoHorarioGenerico(L, S, D, "GL")
            ws.Cells(i, "AI").Value = TextoHorarioGenerico(L, S, D, "CA")
        Else
            ws.Cells(i, "AF").Value = TextoHorario(caso, L, S, D, "EN")
            ws.Cells(i, "AG").Value = TextoHorario(caso, L, S, D, "ES")
            ws.Cells(i, "AH").Value = TextoHorario(caso, L, S, D, "GL")
            ws.Cells(i, "AI").Value = TextoHorario(caso, L, S, D, "CA")
        End If

    Next i

End Sub
