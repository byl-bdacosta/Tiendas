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
'   HELPERS PARA ENCONTRAR COLUMNAS POR ENCABEZADO
'====================================================
Private Function BuscarColumnaEnFila(ws As Worksheet, ByVal fila As Long, ByVal titulo As String) As Long
    Dim ultCol As Long, c As Long
    
    ultCol = ws.Cells(fila, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultCol
        If Trim$(CStr(ws.Cells(fila, c).Value)) = titulo Then
            BuscarColumnaEnFila = c
            Exit Function
        End If
    Next c
End Function

Private Function BuscarColumnaDia(ws As Worksheet, ByVal filaDia As Long, ByVal filaSub As Long, ByVal tituloDia As String) As Long
    Dim ultCol As Long, c As Long
    
    ultCol = ws.Cells(filaDia, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultCol
        If Trim$(CStr(ws.Cells(filaDia, c).Value)) = tituloDia Then
            ' Se asume que en la fila de subencabezado hay "Apertura" justo debajo
            If Trim$(CStr(ws.Cells(filaSub, c).Value)) = "Apertura" Then
                BuscarColumnaDia = c
            Else
                ' Por seguridad, busca "Apertura" hacia la derecha
                Dim c2 As Long
                For c2 = c To ultCol
                    If Trim$(CStr(ws.Cells(filaSub, c2).Value)) = "Apertura" Then
                        BuscarColumnaDia = c2
                        Exit Function
                    End If
                Next c2
            End If
            Exit Function
        End If
    Next c
End Function


'====================================================
'   MACRO PRINCIPAL
'====================================================
Public Sub Horarios()

    Dim ws As Worksheet
    Dim uFila As Long, i As Long
    
    Dim L As HorarioDia, S As HorarioDia, D As HorarioDia
    Dim caso As Integer
    
    Dim filaCabecera1 As Long          ' Fila donde está "COD"
    Dim filaCabecera2 As Long          ' Subcabecera (Apertura/Cierre/Idiomas)
    Dim filaPrimeraDatos As Long
    
    Dim colCOD As Long
    Dim colLIni As Long, colLFin As Long
    Dim colSIni As Long, colSFin As Long
    Dim colDIni As Long, colDFin As Long
    Dim colDom30Ini As Long, colDom30Fin As Long
    
    Dim colEN As Long, colCA As Long, colGL As Long, colES As Long
    
    Set ws = ThisWorkbook.Sheets("HORARIO ESPAÑA")
    
    '------------------------------------------------
    ' Localizar fila de cabecera principal (donde está "COD")
    '------------------------------------------------
    Dim cel As Range
    Set cel = ws.Cells.Find(What:="COD", LookIn:=xlValues, LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If cel Is Nothing Then
        MsgBox "No se ha encontrado la cabecera 'COD' en la hoja HORARIO ESPAÑA.", vbCritical
        Exit Sub
    End If
    
    filaCabecera1 = cel.Row
    colCOD = cel.Column
    filaCabecera2 = filaCabecera1 + 1
    filaPrimeraDatos = filaCabecera2 + 1
    
    '------------------------------------------------
    ' Localizar columnas de días y sus Apertura/Cierre
    ' (usamos la fila de cabecera1 para el nombre del día
    '  y la cabecera2 para 'Apertura'/'Cierre')
    '------------------------------------------------
    colLIni = BuscarColumnaDia(ws, filaCabecera1, filaCabecera2, "Lunes a Viernes")
    colSIni = BuscarColumnaDia(ws, filaCabecera1, filaCabecera2, "Sábado")
    colDIni = BuscarColumnaDia(ws, filaCabecera1, filaCabecera2, "Domingo")
    colDom30Ini = BuscarColumnaDia(ws, filaCabecera1, filaCabecera2, "Domingo 30")
    
    If colLIni = 0 Or colSIni = 0 Or colDIni = 0 Then
        MsgBox "No se han encontrado correctamente las columnas de Lunes/Sábado/Domingo.", vbCritical
        Exit Sub
    End If
    
    colLFin = colLIni + 1
    colSFin = colSIni + 1
    colDFin = colDIni + 1
    If colDom30Ini > 0 Then
        colDom30Fin = colDom30Ini + 1
    End If
    
    '------------------------------------------------
    ' Localizar columnas de idiomas según la fila de subcabecera
    '------------------------------------------------
    colEN = BuscarColumnaEnFila(ws, filaCabecera2, "Inglés")
    colCA = BuscarColumnaEnFila(ws, filaCabecera2, "Catalán")
    colGL = BuscarColumnaEnFila(ws, filaCabecera2, "Gallego")
    colES = BuscarColumnaEnFila(ws, filaCabecera2, "Español")
    
    If colEN = 0 Or colCA = 0 Or colGL = 0 Or colES = 0 Then
        MsgBox "No se han encontrado correctamente las columnas de idiomas (Inglés/Catalán/Gallego/Español).", vbCritical
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Última fila de datos: usamos la columna COD
    '------------------------------------------------
    uFila = ws.Cells(ws.Rows.Count, colCOD).End(xlUp).Row
    
    '------------------------------------------------
    ' Bucle principal
    '------------------------------------------------
    For i = filaPrimeraDatos To uFila
        
        L.Inicio = ws.Cells(i, colLIni).Value
        L.Fin = ws.Cells(i, colLFin).Value
        S.Inicio = ws.Cells(i, colSIni).Value
        S.Fin = ws.Cells(i, colSFin).Value
        D.Inicio = ws.Cells(i, colDIni).Value
        D.Fin = ws.Cells(i, colDFin).Value
        
        '------ Determinar caso ------ 
        caso = ObtenerCaso(L, S, D)
        
        '------ Escribir en Inglés / Catalán / Gallego / Español ------
        If colDom30Ini = 0 Or ws.Cells(i, colDom30Ini).Value = "" Then
            ' Sin horario especial Domingo 30
            ws.Cells(i, colEN).Value = TextoHorario(caso, L, S, D, "EN")
            ws.Cells(i, colCA).Value = TextoHorario(caso, L, S, D, "CA")
            ws.Cells(i, colGL).Value = TextoHorario(caso, L, S, D, "GL")
            ws.Cells(i, colES).Value = TextoHorario(caso, L, S, D, "ES")
        Else
            ' Con horario especial Domingo 30
            ws.Cells(i, colEN).NumberFormat = "@"
            ws.Cells(i, colEN).Value = TextoHorario(caso, L, S, D, "EN") & Chr(10) & _
                                       TextoDomingoEspecial("EN") & FmtHora(ws.Cells(i, colDom30Ini).Value) & " - " & FmtHora(ws.Cells(i, colDom30Fin).Value)
            
            ws.Cells(i, colCA).Value = TextoHorario(caso, L, S, D, "CA") & Chr(10) & _
                                       TextoDomingoEspecial("CA") & FmtHora(ws.Cells(i, colDom30Ini).Value) & " - " & FmtHora(ws.Cells(i, colDom30Fin).Value)
            
            ws.Cells(i, colGL).Value = TextoHorario(caso, L, S, D, "GL") & Chr(10) & _
                                       TextoDomingoEspecial("GL") & FmtHora(ws.Cells(i, colDom30Ini).Value) & " - " & FmtHora(ws.Cells(i, colDom30Fin).Value)
            
            ws.Cells(i, colES).Value = TextoHorario(caso, L, S, D, "ES") & Chr(10) & _
                                       TextoDomingoEspecial("ES") & FmtHora(ws.Cells(i, colDom30Ini).Value) & " - " & FmtHora(ws.Cells(i, colDom30Fin).Value)
        End If
        
    Next i
    
    '------------------------------------------------
    ' Reemplazo de caracteres mal codificados en la hoja
    '------------------------------------------------
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
