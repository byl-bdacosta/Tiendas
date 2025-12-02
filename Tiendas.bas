Option Explicit

'====================================================
'   TIPOS
'====================================================

' Tipo que representa el horario de un día:
' Puede haber hasta 2 turnos (Inicio/Fin y Inicio2/Fin2)
Type HorarioDia
    Inicio As Variant      ' Apertura 1
    Fin As Variant         ' Cierre 1
    Inicio2 As Variant     ' Apertura 2 (opcional)
    Fin2 As Variant        ' Cierre 2 (opcional)
End Type

' Tipo que contiene las columnas detectadas para un día:
' Hasta 2 pares Apertura/Cierre
Type DiaCols
    Ap(1 To 2) As Long     ' Columnas de Apertura
    Ci(1 To 2) As Long     ' Columnas de Cierre
    Num As Integer         ' Número de turnos detectados (1 o 2)
End Type

'====================================================
'   FORMATEADOR DE HORAS
'====================================================
' Convierte un valor a texto "hh:mm", o devuelve vacío si no es hora válida
Private Function FmtHora(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or v = "" Then
        FmtHora = ""
    Else
        FmtHora = Format$(v, "hh:mm")
    End If
End Function

'====================================================
'   COMPARAR DOS DÍAS COMPLETOS (CON 1 O 2 TURNOS)
'====================================================
' Devuelve TRUE si h1 y h2 tienen exactamente los mismos horarios
Private Function HorarioIgual(h1 As HorarioDia, h2 As HorarioDia) As Boolean
    HorarioIgual = (FmtHora(h1.Inicio) = FmtHora(h2.Inicio) And _
                    FmtHora(h1.Fin) = FmtHora(h2.Fin) And _
                    FmtHora(h1.Inicio2) = FmtHora(h2.Inicio2) And _
                    FmtHora(h1.Fin2) = FmtHora(h2.Fin2))
End Function

'====================================================
'   SABER SI UN DÍA ESTÁ VACÍO (CERRADO)
'====================================================
' Devuelve TRUE si no hay ningún horario registrado en el día
Private Function DiaVacio(h As HorarioDia) As Boolean
    DiaVacio = (FmtHora(h.Inicio) = "" And _
                FmtHora(h.Fin) = "" And _
                FmtHora(h.Inicio2) = "" And _
                FmtHora(h.Fin2) = "")
End Function

'====================================================
'   TEXTO PARA UN DÍA (1 O 2 TURNOS)
'====================================================
' Genera un texto como: "10:00 - 14:00 / 16:00 - 20:00"
Private Function TextoDia(h As HorarioDia) As String
    Dim t As String
    
    ' Primer turno
    If FmtHora(h.Inicio) <> "" And FmtHora(h.Fin) <> "" Then
        t = FmtHora(h.Inicio) & " - " & FmtHora(h.Fin)
    End If
    
    ' Segundo turno
    If FmtHora(h.Inicio2) <> "" And FmtHora(h.Fin2) <> "" Then
        If t <> "" Then t = t & " / "
        t = t & FmtHora(h.Inicio2) & " - " & FmtHora(h.Fin2)
    End If
    
    TextoDia = t
End Function

'====================================================
'   DETERMINAR EL CASO (1-2)
'====================================================
' Devuelve:
' 1 → L = S = D
' 2 → L = S ≠ D
' 0 → cualquier otro escenario
Private Function ObtenerCaso(L As HorarioDia, S As HorarioDia, D As HorarioDia) As Integer

    ' Caso 1: los tres son iguales
    If HorarioIgual(L, S) And HorarioIgual(L, D) Then
        ObtenerCaso = 1: Exit Function
    End If
    
    ' Caso 2: L = S pero D es distinto
    If HorarioIgual(L, S) And Not HorarioIgual(L, D) Then
        ObtenerCaso = 2: Exit Function
    End If

    ' Caso 0 por defecto
End Function

'====================================================
'   FORMATO DEL TEXTO (CASOS 0-2)
'====================================================
' Construye la frase completa en el idioma seleccionado
Private Function TextoHorario(caso As Integer, _
                              L As HorarioDia, S As HorarioDia, D As HorarioDia, _
                              idioma As String) As String
                              
    Dim sMonFri As String, sMonSat As String, sMonSun As String
    Dim sSat As String, sSun As String
    Dim sep As String: sep = " | "
    Dim t As String, part As String

    '----- Prefijos por idioma -----
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

    '----- Lógica según el caso -----
    Select Case caso

        Case 1
            ' Caso 1: un único texto "Lun - Dom: ..."
            TextoHorario = sMonSun & TextoDia(L)

        Case 2
            ' Caso 2: "Lun - Sab" + "Dom"
            If DiaVacio(D) Then
                TextoHorario = sMonSat & TextoDia(L)
            Else
                TextoHorario = sMonSat & TextoDia(L) & _
                               sep & sSun & TextoDia(D)
            End If

        Case Else
            ' Caso 0: construir cada día por separado
            t = ""

            If Not DiaVacio(L) Then
                part = sMonFri & TextoDia(L)
                If t <> "" Then t = t & sep
                t = t & part
            End If

            If Not DiaVacio(S) Then
                part = sSat & TextoDia(S)
                If t <> "" Then t = t & sep
                t = t & part
            End If

            If Not DiaVacio(D) Then
                part = sSun & TextoDia(D)
                If t <> "" Then t = t & sep
                t = t & part
            End If

            TextoHorario = t
    End Select
End Function

'====================================================
'   TEXTO DOMINGO 30 DE NOVIEMBRE (CASO 0)
'====================================================
Public Function TextoDomingoEspecial(idioma As String) As String
    Select Case UCase(idioma)
        Case "EN": TextoDomingoEspecial = "Sunday Nov 30: "
        Case "ES": TextoDomingoEspecial = "Domingo 30 Nov: "
        Case "GL": TextoDomingoEspecial = "Domingo 30 Nov: "
        Case "CA": TextoDomingoEspecial = "Diumenge 30 Nov: "
    End Select
End Function

'====================================================
'   HELPERS PARA BUSCAR COLUMNAS
'====================================================

' Busca un título en una sola fila
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

' Busca un texto dentro de un bloque de filas y columnas
Private Function BuscarColumnaTexto(ws As Worksheet, _
                                    ByVal filaIni As Long, ByVal filaFin As Long, _
                                    ByVal titulo As String) As Long
    Dim ultCol As Long, r As Long, c As Long
    
    ultCol = ws.Cells(filaIni, ws.Columns.Count).End(xlToLeft).Column
    
    For r = filaIni To filaFin
        For c = 1 To ultCol
            If Trim$(CStr(ws.Cells(r, c).Value)) = titulo Then
                BuscarColumnaTexto = c
                Exit Function
            End If
        Next c
    Next r
End Function

' Detecta columnas de Apertura y Cierre (hasta 2 turnos) dentro del bloque del día
Private Function DetectarColumnasDia(ws As Worksheet, _
                                     ByVal filaSubIni As Long, ByVal filaSubFin As Long, _
                                     ByVal colIni As Long, ByVal colFin As Long) As DiaCols
    Dim dc As DiaCols
    Dim c As Long, r As Long
    Dim txt As String
    Dim idx As Integer
       
    If colIni = 0 Or colFin = 0 Then
        DetectarColumnasDia = dc
        Exit Function
    End If
    
    ' Recorre cada celda del bloque en busca de "Apertura" o "Cierre"
    For c = colIni To colFin
        For r = filaSubIni To filaSubFin
            txt = Trim$(CStr(ws.Cells(r, c).Value))
            If txt <> "" Then

                ' Columna de apertura
                If StrComp(txt, "Apertura", vbTextCompare) = 0 Then
                    If dc.Num < 2 Then
                        dc.Num = dc.Num + 1
                        dc.Ap(dc.Num) = c
                    End If
                    Exit For

                ' Columna de cierre
                ElseIf StrComp(txt, "Cierre", vbTextCompare) = 0 Then
                    For idx = 1 To 2
                        If dc.Ci(idx) = 0 Then
                            dc.Ci(idx) = c
                            Exit For
                        End If
                    Next idx
                    Exit For   
                End If
            End If
        Next r
    Next c
    
    ' Verificar pares completos Apertura/Cierre
    Dim pares As Integer
    pares = 0
    For idx = 1 To 2
        If dc.Ap(idx) <> 0 And dc.Ci(idx) <> 0 Then pares = pares + 1
    Next idx
    
    dc.Num = pares
    DetectarColumnasDia = dc
End Function


' Lee los valores de horario (Apertura/Cierre) para un día
Private Function LeerHorarioDia(ws As Worksheet, ByVal fila As Long, _
                                cols As DiaCols) As HorarioDia
    Dim h As HorarioDia
    Dim A1 As Variant, C1 As Variant
    Dim A2 As Variant, C2 As Variant
    Dim open1 As Variant, close1 As Variant
    Dim open2 As Variant, close2 As Variant
    
    If cols.Num = 0 Then
        LeerHorarioDia = h
        Exit Function
    End If
    
    ' Leer turno 1
    If cols.Ap(1) <> 0 Then A1 = ws.Cells(fila, cols.Ap(1)).Value
    If cols.Ci(1) <> 0 Then C1 = ws.Cells(fila, cols.Ci(1)).Value
    
    ' Leer turno 2
    If cols.Num >= 2 Then
        If cols.Ap(2) <> 0 Then A2 = ws.Cells(fila, cols.Ap(2)).Value
        If cols.Ci(2) <> 0 Then C2 = ws.Cells(fila, cols.Ci(2)).Value
    End If
    
    ' Si ambos turnos están completos, se usan tal cual
    If FmtHora(A1) <> "" And FmtHora(C1) <> "" And _
       FmtHora(A2) <> "" And FmtHora(C2) <> "" Then
       
        open1 = A1: close1 = C1
        open2 = A2: close2 = C2
        
    Else
        ' Si faltan horarios, se fusionan para generar un horario continuo
        
        ' Determinar apertura más temprana
        If FmtHora(A1) <> "" Then open1 = A1
        If FmtHora(A2) <> "" Then
            If IsEmpty(open1) Or (FmtHora(A2) <> "" And A2 < open1) Then open1 = A2
        End If
        
        ' Determinar cierre más tardío
        If FmtHora(C1) <> "" Then close1 = C1
        If FmtHora(C2) <> "" Then
            If IsEmpty(close1) Or (FmtHora(C2) <> "" And C2 > close1) Then close1 = C2
        End If
    End If
    
    ' Guardar en estructura
    h.Inicio = open1
    h.Fin = close1
    h.Inicio2 = open2
    h.Fin2 = close2
    
    LeerHorarioDia = h
End Function

'====================================================
'   MACRO PRINCIPAL
'====================================================
Public Sub Horarios()

    Dim ws As Worksheet
    Dim uFila As Long, i As Long
    
    Dim L As HorarioDia, S As HorarioDia, D As HorarioDia, D30 As HorarioDia
    Dim caso As Integer
    
    Dim filaCabecera As Long
    Dim filaSubIni As Long, filaSubFin As Long
    Dim filaPrimeraDatos As Long
    
    Dim colCOD As Long
    Dim colLV As Long, colSab As Long, colDom As Long, colDom30 As Long
    Dim colEN As Long, colES As Long, colGL As Long, colCA As Long
    
    Dim colLVIni As Long, colLVFin As Long
    Dim colSabIni As Long, colSabFin As Long
    Dim colDomIni As Long, colDomFin As Long
    Dim colDom30Ini As Long, colDom30Fin As Long
    
    Dim colsLV As DiaCols, colsSab As DiaCols, colsDom As DiaCols, colsDom30 As DiaCols
    
    '------------------------------------------------
    ' Buscar hoja ("Horarios habituales" o "HORARIO ESPAÑA")
    '------------------------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Horarios habituales")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets("HORARIO ESPAÑA")
    End If
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No se ha encontrado la hoja de horarios.", vbCritical
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Localizar la fila donde está "COD"
    '------------------------------------------------
    Dim cel As Range
    Set cel = ws.Cells.Find(What:="COD", LookIn:=xlValues, LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If cel Is Nothing Then
        MsgBox "No se ha encontrado la cabecera 'COD' en la hoja de horarios.", vbCritical
        Exit Sub
    End If

    filaCabecera = cel.Row
    colCOD = cel.Column

    '---- Buscar inicio de subcabeceras ----
    filaSubIni = filaCabecera + 1

    ' Primera fila de datos es el primer COD no vacío
    filaPrimeraDatos = filaCabecera + 1
    Do While IsEmpty(ws.Cells(filaPrimeraDatos, colCOD)) And filaPrimeraDatos < ws.Rows.Count
        filaPrimeraDatos = filaPrimeraDatos + 1
    Loop

    ' Última fila de subcabeceras
    filaSubFin = filaPrimeraDatos - 1

    '------------------------------------------------
    ' Localizar columnas de idiomas
    '------------------------------------------------
    colEN = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Inglés")
    colES = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Español")
    colGL = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Gallego")
    colCA = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Catalán")

    ' Fallback por si no hay acentos
    If colEN = 0 Then colEN = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Ingles")
    If colES = 0 Then colES = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Espanol")
    If colCA = 0 Then colCA = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Catalan")

    If colEN = 0 Or colES = 0 Then
        MsgBox "No se han encontrado correctamente las columnas de idioma.", vbCritical
        Exit Sub
    End If

    '------------------------------------------------
    ' Localizar bloques de días
    '------------------------------------------------
    colLV = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Lunes a Viernes")
    colSab = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Sábado")
    colDom = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Domingo")
    colDom30 = BuscarColumnaTexto(ws, filaCabecera, filaSubFin, "Domingo 30")

    If colLV = 0 Or colSab = 0 Or colDom = 0 Then
        MsgBox "No se han encontrado correctamente las columnas de Lunes/Sábado/Domingo.", vbCritical
        Exit Sub
    End If

    '------------------------------------------------
    ' Definir límites de columnas para cada bloque
    '------------------------------------------------
    colLVIni = colLV
    colSabIni = colSab
    colDomIni = colDom
    If colDom30 > 0 Then colDom30Ini = colDom30
    
    colLVFin = colSabIni - 1
    colSabFin = colDomIni - 1
    
    If colDom30Ini > 0 Then
        colDomFin = colDom30Ini - 1
        colDom30Fin = colEN - 1        ' Antes del bloque de idiomas
    Else
        colDomFin = colEN - 1
    End If

    '------------------------------------------------
    ' Detectar columnas "Apertura/Cierre"
    '------------------------------------------------
    colsLV = DetectarColumnasDia(ws, filaSubIni, filaSubFin, colLVIni, colLVFin)
    colsSab = DetectarColumnasDia(ws, filaSubIni, filaSubFin, colSabIni, colSabFin)
    colsDom = DetectarColumnasDia(ws, filaSubIni, filaSubFin, colDomIni, colDomFin)
    
    If colDom30Ini > 0 Then
        colsDom30 = DetectarColumnasDia(ws, filaSubIni, filaSubFin, colDom30Ini, colDom30Fin)
    End If

    '------------------------------------------------
    ' Última fila de datos
    '------------------------------------------------
    uFila = ws.Cells(ws.Rows.Count, colCOD).End(xlUp).Row

    '------------------------------------------------
    ' Recorrer todas las filas de datos
    '------------------------------------------------
    Dim txtEN As String, txtES As String, txtGL As String, txtCA As String
    
    For i = filaPrimeraDatos To uFila
        
        ' Leer horarios
        L = LeerHorarioDia(ws, i, colsLV)
        S = LeerHorarioDia(ws, i, colsSab)
        D = LeerHorarioDia(ws, i, colsDom)
        
        If colDom30Ini > 0 Then
            D30 = LeerHorarioDia(ws, i, colsDom30)
        Else
            ' Vaciar si no existe
            D30 = Empty
        End If
        
        ' Determinar caso 1-2-0
        caso = ObtenerCaso(L, S, D)
        
        ' Textos básicos
        txtEN = TextoHorario(caso, L, S, D, "EN")
        txtES = TextoHorario(caso, L, S, D, "ES")
        txtGL = TextoHorario(caso, L, S, D, "GL")
        txtCA = TextoHorario(caso, L, S, D, "CA")
        
        ' Añadir Domingo 30 si existe
        If colDom30Ini > 0 And Not DiaVacio(D30) Then
            
            If txtEN <> "" Then txtEN = txtEN & Chr(10)
            If txtES <> "" Then txtES = txtES & Chr(10)
            If txtGL <> "" Then txtGL = txtGL & Chr(10)
            If txtCA <> "" Then txtCA = txtCA & Chr(10)
            
            txtEN = txtEN & TextoDomingoEspecial("EN") & TextoDia(D30)
            txtES = txtES & TextoDomingoEspecial("ES") & TextoDia(D30)
            txtGL = txtGL & TextoDomingoEspecial("GL") & TextoDia(D30)
            txtCA = txtCA & TextoDomingoEspecial("CA") & TextoDia(D30)
        End If
        
        ' Escribir resultado en la hoja
        ws.Cells(i, colEN).NumberFormat = "@": ws.Cells(i, colEN).Value = txtEN
        ws.Cells(i, colES).NumberFormat = "@": ws.Cells(i, colES).Value = txtES
        
        If colGL > 0 Then
            ws.Cells(i, colGL).NumberFormat = "@": ws.Cells(i, colGL).Value = txtGL
        End If
        
        If colCA > 0 Then
            ws.Cells(i, colCA).NumberFormat = "@": ws.Cells(i, colCA).Value = txtCA
        End If
        
    Next i
    
    '------------------------------------------------
    ' Corrección de caracteres mal codificados
    '------------------------------------------------
    With ws.Cells
        .Replace What:="Ã¡", Replacement:="á"
        .Replace What:="Ã©", Replacement:="é"
        .Replace What:="Ã­", Replacement:="í"
        .Replace What:="Ã³", Replacement:="ó"
        .Replace What:="Ãº", Replacement:="ú"

        .Replace What:="Ã", Replacement:="Á"
        .Replace What:="Ã‰", Replacement:="É"
        .Replace What:="Ã", Replacement:="Í"
        .Replace What:="Ã“", Replacement:="Ó"
        .Replace What:="Ãš", Replacement:="Ú"

        .Replace What:="Ã±", Replacement:="ñ"
        .Replace What:="Ã‘", Replacement:="Ñ"

        .Replace What:="Ã¼", Replacement:="ü"
        .Replace What:="Ãœ", Replacement:="Ü"
    End With

End Sub
