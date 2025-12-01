Option Explicit

'====================================================
'   DEFINICIÓN DEL TIPO DE HORARIO POR DÍA
'====================================================
Type HorarioDia
    Inicio As Variant
    Fin As Variant
End Type

'====================================================
'   FORMATEADOR DE HORAS 
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
                    h1.Fin = h2.Fin)
End Function

'====================================================
'   SABER SI UN DÍA ESTÁ VACÍO (CERRADO)
'====================================================
Private Function DiaVacio(h As HorarioDia) As Boolean
    DiaVacio = (FmtHora(h.Inicio) = "" And _
                FmtHora(h.Fin) = "")
End Function

'====================================================
'   DETERMINAR EL CASO (1-2)
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
    
    'Si no coincide ningún caso, se queda en 0 (valor por defecto)

End Function


'====================================================
'   FORMATO DEL TEXTO (CASOS 1-2)
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
            sMonFri = "Mon - Fri: ": sMonSat = "Mon - Sat: ": sMonSun = "Mon - Sun: "
            sSat = "Sat: ": sSun = "Sun: "
        Case "ES"
            sMonFri = "Lun - Vie: ": sMonSat = "Lun - Sáb: ": sMonSun = "Lun - Dom: "
            sSat = "Sáb: ": sSun = "Dom: "
        Case "GL"
            sMonFri = "Lun - Ven: ": sMonSat = "Lun - Sáb: ": sMonSun = "Lun - Dom: "
            sSat = "Sáb: ": sSun = "Dom: "
        Case "CA"
            sMonFri = "Dil - Div: ": sMonSat = "Dil - Dis: ": sMonSun = "Dil - Diu: "
            sSat = "Dis: ": sSun = "Diu: "
        Case Else
            sMonFri = "Mon - Fri: ": sMonSat = "Mon - Sat: ": sMonSun = "Mon - Sun: "
            sSat = "Sat: ": sSun = "Sun: "
    End Select

    '----- CASOS -----
    Select Case caso

        Case 1
            If InStr(L.Inicio, "-") > 0 Then
                TextoHorario = sMonSun & FmtHora(L.Inicio) & " / " & FmtHora(L.Fin)
            Else
                TextoHorario = sMonSun & FmtHora(L.Inicio) & " - " & FmtHora(L.Fin)
            End If

        Case 2
            If DiaVacio(D) Then
                TextoHorario = sMonSat & FmtHora(L.Inicio) & " - " & FmtHora(L.Fin)
            Else
                TextoHorario = sMonSat & FmtHora(L.Inicio) & " - " & FmtHora(L.Fin) & _
                               sep & sSun & FmtHora(D.Inicio) & " - " & FmtHora(D.Fin)
            End If

    End Select
End Function

'====================================================
'   TEXTO DOMINGO 30 DE NOVIEMBRE (CASO 0)
'====================================================
Public Function TextoDomingoEspecial(idioma As String) As String
    Select Case UCase(idioma)
        Case "EN"
            TextoDomingoEspecial = "Sunday Nov 30: "
        Case "ES"
            TextoDomingoEspecial = "Domingo 30 Nov: "
        Case "GL"
            TextoDomingoEspecial = "Domingo 30 Nov: "
        Case "CA"
            TextoDomingoEspecial = "Diumenge 30 Nov: "
    End Select
End Function



'====================================================
'   MACRO PRINCIPAL
'====================================================
Public Sub Horarios()

    Dim ws As Worksheet
    Dim uFila As Long, i As Long
    Dim L As HorarioDia, S As HorarioDia, D As HorarioDia
    Dim caso As Integer

    Set ws = Sheets("HORARIO ESPAÑA")

    uFila = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For i = 4 To uFila

        L.Inicio = ws.Cells(i, "D").Value
        L.Fin = ws.Cells(i, "E").Value
        S.Inicio = ws.Cells(i, "F").Value
        S.Fin = ws.Cells(i, "G").Value
        D.Inicio = ws.Cells(i, "H").Value
        D.Fin = ws.Cells(i, "I").Value


        '------ Determinar caso ------
        caso = ObtenerCaso(L, S, D)

        '------ Escribir en Inglés / Español / Gallego / Catalán ------
        If ws.Cells(i,"J").Value = "" Then
            ws.Cells(i, "P").Value = TextoHorario(caso, L, S, D, "EN")
            ws.Cells(i, "Q").Value = TextoHorario(caso, L, S, D, "CA")
            ws.Cells(i, "R").Value = TextoHorario(caso, L, S, D, "GL")
            ws.Cells(i, "S").Value = TextoHorario(caso, L, S, D, "ES")

        Else
            ws.Cells(i, "P").NumberFormat = "@"
            ws.Cells(i, "P").Value = TextoHorario(caso, L, S, D, "EN") & Chr(10) & TextoDomingoEspecial("EN") & FmtHora(ws.Cells(i, "J").Value) & " - " & FmtHora(ws.Cells(i, "K").Value)
            ws.Cells(i, "Q").Value = TextoHorario(caso, L, S, D, "CA") & Chr(10) &TextoDomingoEspecial("CA") & FmtHora(ws.Cells(i, "J").Value) & " - " & FmtHora(ws.Cells(i, "K").Value)
            ws.Cells(i, "R").Value = TextoHorario(caso, L, S, D, "GL") & Chr(10) & TextoDomingoEspecial("GL") & FmtHora(ws.Cells(i, "J").Value) & " - " & FmtHora(ws.Cells(i, "K").Value)
            ws.Cells(i, "S").Value = TextoHorario(caso, L, S, D, "ES") & Chr(10) & TextoDomingoEspecial("ES") & FmtHora(ws.Cells(i, "J").Value) & " - " & FmtHora(ws.Cells(i, "K").Value)

        End If



    Next i

    ' Reemplazo de caracteres mal codificados en la hoja
    With ws.Cells
        ' Vocales minúsculas
        .Replace What:="Ã¡", Replacement:="á"
        .Replace What:="Ã©", Replacement:="é"
        .Replace What:="Ã­", Replacement:="í"
        .Replace What:="Ã³", Replacement:="ó"
        .Replace What:="Ãº", Replacement:="ú"

        ' Vocales mayúsculas
        .Replace What:="Ã", Replacement:="Á"
        .Replace What:="Ã‰", Replacement:="É"
        .Replace What:="Ã", Replacement:="Í"
        .Replace What:="Ã“", Replacement:="Ó"
        .Replace What:="Ãš", Replacement:="Ú"

        ' Ñ y ñ
        .Replace What:="Ã±", Replacement:="ñ"
        .Replace What:="Ã‘", Replacement:="Ñ"

        ' Ü y ü
        .Replace What:="Ã¼", Replacement:="ü"
        .Replace What:="Ãœ", Replacement:="Ü"
    End With

End Sub
