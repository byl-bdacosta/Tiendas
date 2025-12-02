Option Explicit

'====================================================
'   TIPOS
'====================================================

' Representa el horario de un día (hasta 2 turnos)
Type HorarioDiaInfo
    HoraInicio1 As Variant
    HoraFin1 As Variant
    HoraInicio2 As Variant
    HoraFin2 As Variant
End Type

' Representa las columnas de apertura/cierre detectadas para un día
Type ColumnasDiaInfo
    ColApertura(1 To 2) As Long
    ColCierre(1 To 2) As Long
    NumTurnos As Integer
End Type

'====================================================
'   GESTIÓN DE ERRORES COMÚN
'====================================================
Private Sub GestionarError(ByVal nombreProcedimiento As String)
    MsgBox "Se ha producido un error en: " & nombreProcedimiento & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en macro de horarios"
End Sub

'====================================================
'   FORMATEADOR DE HORAS
'====================================================
Private Function FormatearHora(ByVal valor As Variant) As String
    On Error GoTo ErrHandler
    
    If IsError(valor) Or IsEmpty(valor) Or valor = "" Then
        FormatearHora = ""
    Else
        FormatearHora = Format$(valor, "hh:mm")
    End If
    
    Exit Function
ErrHandler:
    ' En caso de error, devolver vacío para no romper la lógica
    FormatearHora = ""
End Function

'====================================================
'   COMPARAR DOS DÍAS COMPLETOS
'====================================================
Private Function SonHorariosIguales( _
    horario1 As HorarioDiaInfo, _
    horario2 As HorarioDiaInfo) As Boolean
    
    On Error GoTo ErrHandler
    
    SonHorariosIguales = (FormatearHora(horario1.HoraInicio1) = FormatearHora(horario2.HoraInicio1) And _
                          FormatearHora(horario1.HoraFin1) = FormatearHora(horario2.HoraFin1) And _
                          FormatearHora(horario1.HoraInicio2) = FormatearHora(horario2.HoraInicio2) And _
                          FormatearHora(horario1.HoraFin2) = FormatearHora(horario2.HoraFin2))
    Exit Function
ErrHandler:
    SonHorariosIguales = False
End Function

'====================================================
'   SABER SI UN DÍA ESTÁ VACÍO (CERRADO)
'====================================================
Private Function EsDiaVacio(horario As HorarioDiaInfo) As Boolean
    On Error GoTo ErrHandler
    
    EsDiaVacio = (FormatearHora(horario.HoraInicio1) = "" And _
                  FormatearHora(horario.HoraFin1) = "" And _
                  FormatearHora(horario.HoraInicio2) = "" And _
                  FormatearHora(horario.HoraFin2) = "")
    Exit Function
ErrHandler:
    EsDiaVacio = True
End Function

'====================================================
'   TEXTO PARA UN DÍA (1 O 2 TURNOS)
'====================================================
Private Function ObtenerTextoDia(horario As HorarioDiaInfo) As String
    On Error GoTo ErrHandler
    
    Dim texto As String
    
    ' Primer turno
    If FormatearHora(horario.HoraInicio1) <> "" And FormatearHora(horario.HoraFin1) <> "" Then
        texto = FormatearHora(horario.HoraInicio1) & " - " & FormatearHora(horario.HoraFin1)
    End If
    
    ' Segundo turno
    If FormatearHora(horario.HoraInicio2) <> "" And FormatearHora(horario.HoraFin2) <> "" Then
        If texto <> "" Then texto = texto & " / "
        texto = texto & FormatearHora(horario.HoraInicio2) & " - " & FormatearHora(horario.HoraFin2)
    End If
    
    ObtenerTextoDia = texto
    Exit Function
ErrHandler:
    ObtenerTextoDia = ""
End Function

'====================================================
'   DETERMINAR EL CASO (1-2-0)
'====================================================
Private Function ObtenerCasoHorario( _
    horarioLunesViernes As HorarioDiaInfo, _
    horarioSabado As HorarioDiaInfo, _
    horarioDomingo As HorarioDiaInfo) As Integer
    
    On Error GoTo ErrHandler
    
    ' Caso 1: L = S = D
    If SonHorariosIguales(horarioLunesViernes, horarioSabado) And _
       SonHorariosIguales(horarioLunesViernes, horarioDomingo) Then
        ObtenerCasoHorario = 1
        Exit Function
    End If
    
    ' Caso 2: L = S ≠ D
    If SonHorariosIguales(horarioLunesViernes, horarioSabado) And _
       Not SonHorariosIguales(horarioLunesViernes, horarioDomingo) Then
        ObtenerCasoHorario = 2
        Exit Function
    End If
    
    ' Caso 0: genérico
    ObtenerCasoHorario = 0
    Exit Function
ErrHandler:
    ObtenerCasoHorario = 0
End Function

'====================================================
'   FORMATO DEL TEXTO (CASOS 0-2)
'====================================================
Private Function ConstruirTextoHorario( _
    ByVal casoHorario As Integer, _
    horarioLunesViernes As HorarioDiaInfo, _
    horarioSabado As HorarioDiaInfo, _
    horarioDomingo As HorarioDiaInfo, _
    ByVal idioma As String) As String
    
    On Error GoTo ErrHandler
    
    Dim prefijoLunesViernes As String
    Dim prefijoLunesSabado As String
    Dim prefijoLunesDomingo As String
    Dim prefijoSabado As String
    Dim prefijoDomingo As String
    Dim separadorBloques As String
    Dim textoResultado As String
    Dim textoParcial As String
    
    separadorBloques = " | "
    
    ' Prefijos por idioma
    Select Case UCase$(idioma)
        Case "EN"
            prefijoLunesViernes = "Mon - Fri: "
            prefijoLunesSabado = "Mon - Sat: "
            prefijoLunesDomingo = "Mon - Sun: "
            prefijoSabado = "Sat: "
            prefijoDomingo = "Sun: "
        Case "ES"
            prefijoLunesViernes = "Lun - Vie: "
            prefijoLunesSabado = "Lun - Sáb: "
            prefijoLunesDomingo = "Lun - Dom: "
            prefijoSabado = "Sáb: "
            prefijoDomingo = "Dom: "
        Case "GL"
            prefijoLunesViernes = "Lun - Ven: "
            prefijoLunesSabado = "Lun - Sáb: "
            prefijoLunesDomingo = "Lun - Dom: "
            prefijoSabado = "Sáb: "
            prefijoDomingo = "Dom: "
        Case "CA"
            prefijoLunesViernes = "Dil - Div: "
            prefijoLunesSabado = "Dil - Dis: "
            prefijoLunesDomingo = "Dil - Diu: "
            prefijoSabado = "Dis: "
            prefijoDomingo = "Diu: "
        Case Else
            prefijoLunesViernes = "Mon - Fri: "
            prefijoLunesSabado = "Mon - Sat: "
            prefijoLunesDomingo = "Mon - Sun: "
            prefijoSabado = "Sat: "
            prefijoDomingo = "Sun: "
    End Select
    
    Select Case casoHorario
        Case 1
            ' Caso 1: L = S = D → "Lun - Dom: ..."
            ConstruirTextoHorario = prefijoLunesDomingo & ObtenerTextoDia(horarioLunesViernes)
        
        Case 2
            ' Caso 2: L = S ≠ D
            If EsDiaVacio(horarioDomingo) Then
                ConstruirTextoHorario = prefijoLunesSabado & ObtenerTextoDia(horarioLunesViernes)
            Else
                ConstruirTextoHorario = prefijoLunesSabado & ObtenerTextoDia(horarioLunesViernes) & _
                                        separadorBloques & prefijoDomingo & ObtenerTextoDia(horarioDomingo)
            End If
        
        Case Else
            ' Caso 0: genérico
            textoResultado = ""
            
            If Not EsDiaVacio(horarioLunesViernes) Then
                textoParcial = prefijoLunesViernes & ObtenerTextoDia(horarioLunesViernes)
                If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
                textoResultado = textoResultado & textoParcial
            End If
            
            If Not EsDiaVacio(horarioSabado) Then
                textoParcial = prefijoSabado & ObtenerTextoDia(horarioSabado)
                If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
                textoResultado = textoResultado & textoParcial
            End If
            
            If Not EsDiaVacio(horarioDomingo) Then
                textoParcial = prefijoDomingo & ObtenerTextoDia(horarioDomingo)
                If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
                textoResultado = textoResultado & textoParcial
            End If
            
            ConstruirTextoHorario = textoResultado
    End Select
    
    Exit Function
ErrHandler:
    ConstruirTextoHorario = ""
End Function

'====================================================
'   TEXTO DOMINGO ESPECIAL (DOMINGO 30)
'====================================================
Public Function TextoDomingoEspecial(ByVal idioma As String) As String
    On Error GoTo ErrHandler
    
    Select Case UCase$(idioma)
        Case "EN": TextoDomingoEspecial = "Sunday Nov 30: "
        Case "ES": TextoDomingoEspecial = "Domingo 30 Nov: "
        Case "GL": TextoDomingoEspecial = "Domingo 30 Nov: "
        Case "CA": TextoDomingoEspecial = "Diumenge 30 Nov: "
        Case Else: TextoDomingoEspecial = "Sunday Nov 30: "
    End Select
    
    Exit Function
ErrHandler:
    TextoDomingoEspecial = ""
End Function

'====================================================
'   TEXTO DÍAS ESPECIALES 6-7-8 DIC
'====================================================
Private Function TextoDiaEspecialDic( _
    ByVal idioma As String, _
    ByVal codigoDia As String) As String
    ' codigoDia: "SAB6", "DOM7", "LUN8"
    
    On Error GoTo ErrHandler
    
    Select Case UCase$(idioma)
        Case "ES"
            Select Case UCase$(codigoDia)
                Case "SAB6": TextoDiaEspecialDic = "Sábado 6 de Dic: "
                Case "DOM7": TextoDiaEspecialDic = "Domingo 7 de Dic: "
                Case "LUN8": TextoDiaEspecialDic = "Lunes 8 de Dic: "
            End Select
        
        Case "EN"
            Select Case UCase$(codigoDia)
                Case "SAB6": TextoDiaEspecialDic = "Sat Dec 6: "
                Case "DOM7": TextoDiaEspecialDic = "Sun Dec 7: "
                Case "LUN8": TextoDiaEspecialDic = "Mon Dec 8: "
            End Select
        
        Case "GL"
            Select Case UCase$(codigoDia)
                Case "SAB6": TextoDiaEspecialDic = "Sábado 6 de Dec: "
                Case "DOM7": TextoDiaEspecialDic = "Domingo 7 de Dec: "
                Case "LUN8": TextoDiaEspecialDic = "Luns 8 de Dec: "
            End Select
        
        Case "CA"
            Select Case UCase$(codigoDia)
                Case "SAB6": TextoDiaEspecialDic = "Dissabte 6 de Des: "
                Case "DOM7": TextoDiaEspecialDic = "Diumenge 7 de Des: "
                Case "LUN8": TextoDiaEspecialDic = "Dilluns 8 de Des: "
            End Select
        
        Case Else
            Select Case UCase$(codigoDia)
                Case "SAB6": TextoDiaEspecialDic = "Sat Dec 6: "
                Case "DOM7": TextoDiaEspecialDic = "Sun Dec 7: "
                Case "LUN8": TextoDiaEspecialDic = "Mon Dec 8: "
            End Select
    End Select
    
    Exit Function
ErrHandler:
    TextoDiaEspecialDic = ""
End Function

'====================================================
'   HELPERS PARA BUSCAR COLUMNAS
'====================================================
Private Function BuscarColumnaEnFilaPorTitulo( _
    hoja As Worksheet, _
    ByVal fila As Long, _
    ByVal titulo As String) As Long
    
    On Error GoTo ErrHandler
    
    Dim ultimaColumna As Long
    Dim col As Long
    
    ultimaColumna = hoja.Cells(fila, hoja.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To ultimaColumna
        If Trim$(CStr(hoja.Cells(fila, col).Value)) = titulo Then
            BuscarColumnaEnFilaPorTitulo = col
            Exit Function
        End If
    Next col
    
    BuscarColumnaEnFilaPorTitulo = 0
    Exit Function
ErrHandler:
    BuscarColumnaEnFilaPorTitulo = 0
End Function

Private Function BuscarColumnaPorTextoEnBloque( _
    hoja As Worksheet, _
    ByVal filaInicio As Long, _
    ByVal filaFin As Long, _
    ByVal titulo As String) As Long
    
    On Error GoTo ErrHandler
    
    Dim ultimaColumna As Long
    Dim fila As Long
    Dim col As Long
    
    ultimaColumna = hoja.Cells(filaInicio, hoja.Columns.Count).End(xlToLeft).Column
    
    For fila = filaInicio To filaFin
        For col = 1 To ultimaColumna
            If Trim$(CStr(hoja.Cells(fila, col).Value)) = titulo Then
                BuscarColumnaPorTextoEnBloque = col
                Exit Function
            End If
        Next col
    Next fila
    
    BuscarColumnaPorTextoEnBloque = 0
    Exit Function
ErrHandler:
    BuscarColumnaPorTextoEnBloque = 0
End Function

' Devuelve la siguiente columna de inicio de bloque a la derecha de colActual
Private Function SiguienteInicio(ByVal colActual As Long, ParamArray candidatos() As Variant) As Long
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim c As Long
    Dim menor As Long
    
    menor = 0
    For i = LBound(candidatos) To UBound(candidatos)
        c = CLng(candidatos(i))
        If c <> 0 And c > colActual Then
            If menor = 0 Or c < menor Then menor = c
        End If
    Next i
    
    SiguienteInicio = menor
    Exit Function
ErrHandler:
    SiguienteInicio = 0
End Function

'====================================================
'   DETECTAR COLUMNAS APERTURA/CIERRE (ROBUSTO)
'====================================================
Private Function DetectarColumnasAperturaCierreDia( _
    hoja As Worksheet, _
    ByVal filaSubcabeceraInicio As Long, _
    ByVal filaSubcabeceraFin As Long, _
    ByVal colInicio As Long, _
    ByVal colFin As Long) As ColumnasDiaInfo
    
    On Error GoTo ErrHandler
    
    Dim infoColumnas As ColumnasDiaInfo
    Dim col As Long, fila As Long
    Dim textoCelda As String
    Dim indice As Integer
    Dim numParesCompletos As Integer
    
    If colInicio = 0 Or colFin = 0 Then
        DetectarColumnasAperturaCierreDia = infoColumnas
        Exit Function
    End If
    
    For col = colInicio To colFin
        For fila = filaSubcabeceraInicio To filaSubcabeceraFin
            textoCelda = LCase$(Trim$(CStr(hoja.Cells(fila, col).Value)))
            
            If textoCelda <> "" Then
                ' Cualquier cosa que contenga "apert" cuenta como apertura
                If InStr(1, textoCelda, "apert", vbTextCompare) > 0 Then
                    If infoColumnas.NumTurnos < 2 Then
                        infoColumnas.NumTurnos = infoColumnas.NumTurnos + 1
                        infoColumnas.ColApertura(infoColumnas.NumTurnos) = col
                    End If
                    Exit For
                ' Cualquier cosa que contenga "cierre" cuenta como cierre
                ElseIf InStr(1, textoCelda, "cierre", vbTextCompare) > 0 Then
                    For indice = 1 To 2
                        If infoColumnas.ColCierre(indice) = 0 Then
                            infoColumnas.ColCierre(indice) = col
                            Exit For
                        End If
                    Next indice
                    Exit For
                End If
            End If
        Next fila
    Next col
    
    ' Contar pares Apertura/Cierre completos
    numParesCompletos = 0
    For indice = 1 To 2
        If infoColumnas.ColApertura(indice) <> 0 And infoColumnas.ColCierre(indice) <> 0 Then
            numParesCompletos = numParesCompletos + 1
        End If
    Next indice
    
    infoColumnas.NumTurnos = numParesCompletos
    DetectarColumnasAperturaCierreDia = infoColumnas
    Exit Function
ErrHandler:
    ' Devuelve estructura vacía
    DetectarColumnasAperturaCierreDia = infoColumnas
End Function

'====================================================
'   LECTURA DE HORARIO DE UN DÍA
'====================================================
Private Function LeerHorarioDeDia( _
    hoja As Worksheet, _
    ByVal fila As Long, _
    columnasDia As ColumnasDiaInfo) As HorarioDiaInfo
    
    On Error GoTo ErrHandler
    
    Dim horario As HorarioDiaInfo
    Dim apertura1 As Variant, cierre1 As Variant
    Dim apertura2 As Variant, cierre2 As Variant
    Dim horaInicioUnica As Variant, horaFinUnica As Variant
    Dim horaInicio2Turno As Variant, horaFin2Turno As Variant
    
    If columnasDia.NumTurnos = 0 Then
        LeerHorarioDeDia = horario
        Exit Function
    End If
    
    ' Turno 1
    If columnasDia.ColApertura(1) <> 0 Then apertura1 = hoja.Cells(fila, columnasDia.ColApertura(1)).Value
    If columnasDia.ColCierre(1) <> 0 Then cierre1 = hoja.Cells(fila, columnasDia.ColCierre(1)).Value
    
    ' Turno 2
    If columnasDia.NumTurnos >= 2 Then
        If columnasDia.ColApertura(2) <> 0 Then apertura2 = hoja.Cells(fila, columnasDia.ColApertura(2)).Value
        If columnasDia.ColCierre(2) <> 0 Then cierre2 = hoja.Cells(fila, columnasDia.ColCierre(2)).Value
    End If
    
    ' Si hay 2 turnos completos se respetan tal cual
    If FormatearHora(apertura1) <> "" And FormatearHora(cierre1) <> "" And _
       FormatearHora(apertura2) <> "" And FormatearHora(cierre2) <> "" Then
       
        horaInicioUnica = apertura1: horaFinUnica = cierre1
        horaInicio2Turno = apertura2: horaFin2Turno = cierre2
    
    Else
        ' Si hay huecos se fusionan para un horario continuo
        
        ' Apertura más temprana
        If FormatearHora(apertura1) <> "" Then horaInicioUnica = apertura1
        If FormatearHora(apertura2) <> "" Then
            If IsEmpty(horaInicioUnica) Or (FormatearHora(apertura2) <> "" And apertura2 < horaInicioUnica) Then
                horaInicioUnica = apertura2
            End If
        End If
        
        ' Cierre más tardío
        If FormatearHora(cierre1) <> "" Then horaFinUnica = cierre1
        If FormatearHora(cierre2) <> "" Then
            If IsEmpty(horaFinUnica) Or (FormatearHora(cierre2) <> "" And cierre2 > horaFinUnica) Then
                horaFinUnica = cierre2
            End If
        End If
    End If
    
    ' Guardar en estructura
    horario.HoraInicio1 = horaInicioUnica
    horario.HoraFin1 = horaFinUnica
    horario.HoraInicio2 = horaInicio2Turno
    horario.HoraFin2 = horaFin2Turno
    
    LeerHorarioDeDia = horario
    Exit Function
ErrHandler:
    ' Devuelve lo que haya (vacío por defecto)
    LeerHorarioDeDia = horario
End Function

'====================================================
'   MACRO PRINCIPAL
'====================================================
Public Sub Horarios()

    Const NOMBRE_PROC As String = "Horarios"
    On Error GoTo ErrHandler
    
    Dim hojaHorarios As Worksheet
    Dim ultimaFilaDatos As Long
    Dim filaActual As Long
    
    Dim horarioLunesViernes As HorarioDiaInfo
    Dim horarioSabado As HorarioDiaInfo
    Dim horarioDomingo As HorarioDiaInfo
    Dim horarioDomingoEspecial As HorarioDiaInfo
    
    ' Días especiales 6-7-8 Dic
    Dim horarioSabado6 As HorarioDiaInfo
    Dim horarioDomingo7Dic As HorarioDiaInfo
    Dim horarioLunes8 As HorarioDiaInfo
    
    Dim casoHorario As Integer
    
    Dim filaCabeceraPrincipal As Long
    Dim filaSubcabeceraInicio As Long
    Dim filaSubcabeceraFin As Long
    Dim filaPrimeraDatos As Long
    
    Dim colCOD As Long
    Dim colLunesViernes As Long
    Dim colSabado As Long
    Dim colDomingo As Long
    Dim colDomingo30 As Long
    Dim colSabado6 As Long
    Dim colDomingo7Dic As Long
    Dim colLunes8 As Long
    
    Dim colEN As Long, colES As Long, colGL As Long, colCA As Long
    
    Dim colLunesViernesInicio As Long, colLunesViernesFin As Long
    Dim colSabadoInicio As Long, colSabadoFin As Long
    Dim colDomingoInicio As Long, colDomingoFin As Long
    Dim colDomingo30Inicio As Long, colDomingo30Fin As Long
    Dim colSabado6Inicio As Long, colSabado6Fin As Long
    Dim colDomingo7Inicio As Long, colDomingo7Fin As Long
    Dim colLunes8Inicio As Long, colLunes8Fin As Long
    
    Dim columnasLunesViernes As ColumnasDiaInfo
    Dim columnasSabado As ColumnasDiaInfo
    Dim columnasDomingo As ColumnasDiaInfo
    Dim columnasDomingo30 As ColumnasDiaInfo
    Dim columnasSabado6 As ColumnasDiaInfo
    Dim columnasDomingo7Dic As ColumnasDiaInfo
    Dim columnasLunes8 As ColumnasDiaInfo
    
    Dim textoEN As String, textoES As String, textoGL As String, textoCA As String
    Dim celdaCabeceraCOD As Range
    Dim celdaTmp As Range
    
    ' acumuladores de días especiales (6-7-8 dic)
    Dim extrasEN As String, extrasES As String, extrasGL As String, extrasCA As String
    
    '------------------------------------------------
    ' Localizar hoja
    '------------------------------------------------
    On Error Resume Next
    Set hojaHorarios = ThisWorkbook.Worksheets("Horarios habituales")
    If hojaHorarios Is Nothing Then
        Set hojaHorarios = ThisWorkbook.Worksheets("HORARIO ESPAÑA")
    End If
    On Error GoTo ErrHandler
    
    If hojaHorarios Is Nothing Then
        MsgBox "No se ha encontrado la hoja de horarios (""Horarios habituales"" o ""HORARIO ESPAÑA"").", _
               vbCritical, "Horarios"
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Localizar cabecera COD
    '------------------------------------------------
    Set celdaCabeceraCOD = hojaHorarios.Cells.Find(What:="COD", LookIn:=xlValues, LookAt:=xlWhole, _
                                                   SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If celdaCabeceraCOD Is Nothing Then
        MsgBox "No se ha encontrado la cabecera 'COD' en la hoja de horarios.", vbCritical, "Horarios"
        Exit Sub
    End If
    
    filaCabeceraPrincipal = celdaCabeceraCOD.Row
    colCOD = celdaCabeceraCOD.Column
    
    ' Subcabeceras
    filaSubcabeceraInicio = filaCabeceraPrincipal + 1
    
    ' Primera fila de datos: primer COD no vacío
    filaPrimeraDatos = filaCabeceraPrincipal + 1
    Do While IsEmpty(hojaHorarios.Cells(filaPrimeraDatos, colCOD)) And _
             filaPrimeraDatos < hojaHorarios.Rows.Count
        filaPrimeraDatos = filaPrimeraDatos + 1
    Loop
    
    filaSubcabeceraFin = filaPrimeraDatos - 1
    
    '------------------------------------------------
    ' Columnas de idiomas (búsqueda robusta)
    '------------------------------------------------
    Set celdaTmp = hojaHorarios.Cells.Find(What:="Inglés", LookIn:=xlValues, LookAt:=xlWhole, _
                                           SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If celdaTmp Is Nothing Then
        Set celdaTmp = hojaHorarios.Cells.Find(What:="Ingles", LookIn:=xlValues, LookAt:=xlWhole, _
                                               SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End If
    If Not celdaTmp Is Nothing Then colEN = celdaTmp.Column
    
    Set celdaTmp = hojaHorarios.Cells.Find(What:="Español", LookIn:=xlValues, LookAt:=xlWhole, _
                                           SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If celdaTmp Is Nothing Then
        Set celdaTmp = hojaHorarios.Cells.Find(What:="Espanol", LookIn:=xlValues, LookAt:=xlWhole, _
                                               SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End If
    If Not celdaTmp Is Nothing Then colES = celdaTmp.Column
    
    Set celdaTmp = hojaHorarios.Cells.Find(What:="Gallego", LookIn:=xlValues, LookAt:=xlWhole, _
                                           SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not celdaTmp Is Nothing Then colGL = celdaTmp.Column
    
    Set celdaTmp = hojaHorarios.Cells.Find(What:="Catalán", LookIn:=xlValues, LookAt:=xlWhole, _
                                           SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If celdaTmp Is Nothing Then
        Set celdaTmp = hojaHorarios.Cells.Find(What:="Catalan", LookIn:=xlValues, LookAt:=xlWhole, _
                                               SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End If
    If Not celdaTmp Is Nothing Then colCA = celdaTmp.Column
    
    If colEN = 0 Or colES = 0 Then
        MsgBox "No se han encontrado correctamente las columnas de idioma (al menos Inglés y Español).", _
               vbCritical, "Horarios"
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Bloques de días
    '------------------------------------------------
    colLunesViernes = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Lunes a Viernes")
    colSabado = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Sábado")
    
    ' Domingo habitual puede no existir
    colDomingo = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Domingo")
    
    colDomingo30 = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Domingo 30")
    
    ' Días especiales de diciembre 2025
    colSabado6 = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Sábado 06-12")
    If colSabado6 = 0 Then
        colSabado6 = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Sabado 06-12")
        If colSabado6 = 0 Then
            colSabado6 = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Sábado 6")
            If colSabado6 = 0 Then colSabado6 = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Sabado 6")
        End If
    End If
    
    colDomingo7Dic = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Domingo 07-12")
    If colDomingo7Dic = 0 Then
        colDomingo7Dic = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Domingo 7")
    End If
    
    colLunes8 = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Lunes 08-12")
    If colLunes8 = 0 Then
        colLunes8 = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Lunes 8")
    End If
    
    ' Lunes y Sábado habituales son obligatorios
    If colLunesViernes = 0 Or colSabado = 0 Then
        MsgBox "No se han encontrado correctamente las columnas de Lunes/Sábado.", _
               vbCritical, "Horarios"
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Rangos de columnas para cada bloque
    '------------------------------------------------
    colLunesViernesInicio = colLunesViernes
    colLunesViernesFin = SiguienteInicio(colLunesViernesInicio, _
                                         colSabado, colDomingo, colSabado6, colDomingo7Dic, colLunes8, colDomingo30, colEN) - 1
    
    colSabadoInicio = colSabado
    colSabadoFin = SiguienteInicio(colSabadoInicio, _
                                   colDomingo, colSabado6, colDomingo7Dic, colLunes8, colDomingo30, colEN) - 1
    
    If colDomingo > 0 Then
        colDomingoInicio = colDomingo
        colDomingoFin = SiguienteInicio(colDomingoInicio, _
                                        colSabado6, colDomingo7Dic, colLunes8, colDomingo30, colEN) - 1
    Else
        colDomingoInicio = 0
        colDomingoFin = 0
    End If
    
    If colSabado6 > 0 Then
        colSabado6Inicio = colSabado6
        colSabado6Fin = SiguienteInicio(colSabado6Inicio, _
                                        colDomingo7Dic, colLunes8, colDomingo30, colEN) - 1
    End If
    
    If colDomingo7Dic > 0 Then
        colDomingo7Inicio = colDomingo7Dic
        colDomingo7Fin = SiguienteInicio(colDomingo7Inicio, _
                                         colLunes8, colDomingo30, colEN) - 1
    End If
    
    If colLunes8 > 0 Then
        colLunes8Inicio = colLunes8
        colLunes8Fin = SiguienteInicio(colLunes8Inicio, _
                                       colDomingo30, colEN) - 1
    End If
    
    If colDomingo30 > 0 Then
        colDomingo30Inicio = colDomingo30
        colDomingo30Fin = SiguienteInicio(colDomingo30Inicio, _
                                          colEN) - 1
    End If
    
    '------------------------------------------------
    ' Detectar columnas de Apertura/Cierre por bloque
    '------------------------------------------------
    columnasLunesViernes = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                             colLunesViernesInicio, colLunesViernesFin)
    columnasSabado = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                       colSabadoInicio, colSabadoFin)
    columnasDomingo = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                        colDomingoInicio, colDomingoFin)
    If colDomingo30Inicio > 0 Then
        columnasDomingo30 = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                              colDomingo30Inicio, colDomingo30Fin)
    End If
    
    If colSabado6Inicio > 0 Then
        columnasSabado6 = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                            colSabado6Inicio, colSabado6Fin)
    End If
    
    If colDomingo7Inicio > 0 Then
        columnasDomingo7Dic = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                                colDomingo7Inicio, colDomingo7Fin)
    End If
    
    If colLunes8Inicio > 0 Then
        columnasLunes8 = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                           colLunes8Inicio, colLunes8Fin)
    End If
    
    '------------------------------------------------
    ' Última fila de datos (por COD)
    '------------------------------------------------
    ultimaFilaDatos = hojaHorarios.Cells(hojaHorarios.Rows.Count, colCOD).End(xlUp).Row
    
    '------------------------------------------------
    ' Bucle principal por filas de datos
    '------------------------------------------------
    For filaActual = filaPrimeraDatos To ultimaFilaDatos
        
        ' Reiniciar acumuladores de días especiales para ESTA fila
        extrasEN = ""
        extrasES = ""
        extrasGL = ""
        extrasCA = ""
        
        ' Leer horarios por día
        horarioLunesViernes = LeerHorarioDeDia(hojaHorarios, filaActual, columnasLunesViernes)
        horarioSabado = LeerHorarioDeDia(hojaHorarios, filaActual, columnasSabado)
        horarioDomingo = LeerHorarioDeDia(hojaHorarios, filaActual, columnasDomingo)
        
        If colDomingo30Inicio > 0 Then
            horarioDomingoEspecial = LeerHorarioDeDia(hojaHorarios, filaActual, columnasDomingo30)
        Else
            With horarioDomingoEspecial
                .HoraInicio1 = Empty: .HoraFin1 = Empty
                .HoraInicio2 = Empty: .HoraFin2 = Empty
            End With
        End If
        
        If colSabado6Inicio > 0 Then
            horarioSabado6 = LeerHorarioDeDia(hojaHorarios, filaActual, columnasSabado6)
        Else
            With horarioSabado6
                .HoraInicio1 = Empty: .HoraFin1 = Empty
                .HoraInicio2 = Empty: .HoraFin2 = Empty
            End With
        End If
        
        If colDomingo7Inicio > 0 Then
            horarioDomingo7Dic = LeerHorarioDeDia(hojaHorarios, filaActual, columnasDomingo7Dic)
        Else
            With horarioDomingo7Dic
                .HoraInicio1 = Empty: .HoraFin1 = Empty
                .HoraInicio2 = Empty: .HoraFin2 = Empty
            End With
        End If
        
        If colLunes8Inicio > 0 Then
            horarioLunes8 = LeerHorarioDeDia(hojaHorarios, filaActual, columnasLunes8)
        Else
            With horarioLunes8
                .HoraInicio1 = Empty: .HoraFin1 = Empty
                .HoraInicio2 = Empty: .HoraFin2 = Empty
            End With
        End If
        
        ' Determinar caso
        casoHorario = ObtenerCasoHorario(horarioLunesViernes, horarioSabado, horarioDomingo)
        
        ' Textos base
        textoEN = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "EN")
        textoES = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "ES")
        textoGL = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "GL")
        textoCA = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "CA")
        
        ' Domingo 30 especial
        If colDomingo30Inicio > 0 And Not EsDiaVacio(horarioDomingoEspecial) Then
            
            If textoEN <> "" Then textoEN = textoEN & Chr(10)
            If textoES <> "" Then textoES = textoES & Chr(10)
            If textoGL <> "" Then textoGL = textoGL & Chr(10)
            If textoCA <> "" Then textoCA = textoCA & Chr(10)
            
            textoEN = textoEN & TextoDomingoEspecial("EN") & ObtenerTextoDia(horarioDomingoEspecial)
            textoES = textoES & TextoDomingoEspecial("ES") & ObtenerTextoDia(horarioDomingoEspecial)
            textoGL = textoGL & TextoDomingoEspecial("GL") & ObtenerTextoDia(horarioDomingoEspecial)
            textoCA = textoCA & TextoDomingoEspecial("CA") & ObtenerTextoDia(horarioDomingoEspecial)
        End If
        
        '------------------------------------------------
        ' DÍAS ESPECIALES 6-7-8 DIC EN BLOQUE ÚNICO
        '------------------------------------------------
        ' Sábado 6
        If Not EsDiaVacio(horarioSabado6) Then
            extrasEN = extrasEN & TextoDiaEspecialDic("EN", "SAB6") & ObtenerTextoDia(horarioSabado6)
            extrasES = extrasES & TextoDiaEspecialDic("ES", "SAB6") & ObtenerTextoDia(horarioSabado6)
            extrasGL = extrasGL & TextoDiaEspecialDic("GL", "SAB6") & ObtenerTextoDia(horarioSabado6)
            extrasCA = extrasCA & TextoDiaEspecialDic("CA", "SAB6") & ObtenerTextoDia(horarioSabado6)
        End If
        
        ' Domingo 7
        If Not EsDiaVacio(horarioDomingo7Dic) Then
            If extrasEN <> "" Then extrasEN = extrasEN & " | "
            If extrasES <> "" Then extrasES = extrasES & " | "
            If extrasGL <> "" Then extrasGL = extrasGL & " | "
            If extrasCA <> "" Then extrasCA = extrasCA & " | "
            
            extrasEN = extrasEN & TextoDiaEspecialDic("EN", "DOM7") & ObtenerTextoDia(horarioDomingo7Dic)
            extrasES = extrasES & TextoDiaEspecialDic("ES", "DOM7") & ObtenerTextoDia(horarioDomingo7Dic)
            extrasGL = extrasGL & TextoDiaEspecialDic("GL", "DOM7") & ObtenerTextoDia(horarioDomingo7Dic)
            extrasCA = extrasCA & TextoDiaEspecialDic("CA", "DOM7") & ObtenerTextoDia(horarioDomingo7Dic)
        End If
        
        ' Lunes 8
        If Not EsDiaVacio(horarioLunes8) Then
            If extrasEN <> "" Then extrasEN = extrasEN & " | "
            If extrasES <> "" Then extrasES = extrasES & " | "
            If extrasGL <> "" Then extrasGL = extrasGL & " | "
            If extrasCA <> "" Then extrasCA = extrasCA & " | "
            
            extrasEN = extrasEN & TextoDiaEspecialDic("EN", "LUN8") & ObtenerTextoDia(horarioLunes8)
            extrasES = extrasES & TextoDiaEspecialDic("ES", "LUN8") & ObtenerTextoDia(horarioLunes8)
            extrasGL = extrasGL & TextoDiaEspecialDic("GL", "LUN8") & ObtenerTextoDia(horarioLunes8)
            extrasCA = extrasCA & TextoDiaEspecialDic("CA", "LUN8") & ObtenerTextoDia(horarioLunes8)
        End If
        
        ' Si hay algún especial, añadirlo en una nueva línea
        If extrasEN <> "" Then
            If textoEN <> "" Then textoEN = textoEN & Chr(10)
            textoEN = textoEN & extrasEN
        End If
        
        If extrasES <> "" Then
            If textoES <> "" Then textoES = textoES & Chr(10)
            textoES = textoES & extrasES
        End If
        
        If extrasGL <> "" Then
            If textoGL <> "" Then textoGL = textoGL & Chr(10)
            textoGL = textoGL & extrasGL
        End If
        
        If extrasCA <> "" Then
            If textoCA <> "" Then textoCA = textoCA & Chr(10)
            textoCA = textoCA & extrasCA
        End If
        
        '------------------------------------------------
        ' Escribir en celdas
        '------------------------------------------------
        hojaHorarios.Cells(filaActual, colEN).NumberFormat = "@"
        hojaHorarios.Cells(filaActual, colEN).Value = textoEN
        
        hojaHorarios.Cells(filaActual, colES).NumberFormat = "@"
        hojaHorarios.Cells(filaActual, colES).Value = textoES
        
        If colGL > 0 Then
            hojaHorarios.Cells(filaActual, colGL).NumberFormat = "@"
            hojaHorarios.Cells(filaActual, colGL).Value = textoGL
        End If
        
        If colCA > 0 Then
            hojaHorarios.Cells(filaActual, colCA).NumberFormat = "@"
            hojaHorarios.Cells(filaActual, colCA).Value = textoCA
        End If
    Next filaActual
    
    '------------------------------------------------
    ' Corrección de caracteres mal codificados
    '------------------------------------------------
    With hojaHorarios.Cells
        ' Vocales minúsculas
        .Replace What:="Ã¡", Replacement:="á", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã©", Replacement:="é", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã­", Replacement:="í", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã³", Replacement:="ó", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ãº", Replacement:="ú", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        
        ' Vocales mayúsculas
        .Replace What:="Ã", Replacement:="Á", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã‰", Replacement:="É", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã", Replacement:="Í", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã“", Replacement:="Ó", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ãš", Replacement:="Ú", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        
        ' Ñ y ñ
        .Replace What:="Ã±", Replacement:="ñ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã‘", Replacement:="Ñ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        
        ' Ü y ü
        .Replace What:="Ã¼", Replacement:="ü", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ãœ", Replacement:="Ü", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    End With
    
    Exit Sub
ErrHandler:
    GestionarError NOMBRE_PROC
End Sub
