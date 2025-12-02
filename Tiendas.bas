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
            textoCelda = Trim$(CStr(hoja.Cells(fila, col).Value))
            
            If textoCelda <> "" Then
                If StrComp(textoCelda, "Apertura", vbTextCompare) = 0 Then
                    If infoColumnas.NumTurnos < 2 Then
                        infoColumnas.NumTurnos = infoColumnas.NumTurnos + 1
                        infoColumnas.ColApertura(infoColumnas.NumTurnos) = col
                    End If
                    Exit For
                ElseIf StrComp(textoCelda, "Cierre", vbTextCompare) = 0 Then
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
    Dim colEN As Long, colES As Long, colGL As Long, colCA As Long
    
    Dim colLunesViernesInicio As Long, colLunesViernesFin As Long
    Dim colSabadoInicio As Long, colSabadoFin As Long
    Dim colDomingoInicio As Long, colDomingoFin As Long
    Dim colDomingo30Inicio As Long, colDomingo30Fin As Long
    
    Dim columnasLunesViernes As ColumnasDiaInfo
    Dim columnasSabado As ColumnasDiaInfo
    Dim columnasDomingo As ColumnasDiaInfo
    Dim columnasDomingo30 As ColumnasDiaInfo
    
    Dim textoEN As String, textoES As String, textoGL As String, textoCA As String
    Dim celdaCabeceraCOD As Range
    
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
    ' Columnas de idiomas
    '------------------------------------------------
    colEN = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Inglés")
    colES = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Español")
    colGL = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Gallego")
    colCA = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Catalán")
    
    ' Fallback sin acentos
    If colEN = 0 Then colEN = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Ingles")
    If colES = 0 Then colES = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Espanol")
    If colCA = 0 Then colCA = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Catalan")
    
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
    colDomingo = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Domingo")
    colDomingo30 = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Domingo 30")
    
    If colLunesViernes = 0 Or colSabado = 0 Or colDomingo = 0 Then
        MsgBox "No se han encontrado correctamente las columnas de Lunes/Sábado/Domingo.", _
               vbCritical, "Horarios"
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Rangos de columnas para cada bloque
    '------------------------------------------------
    colLunesViernesInicio = colLunesViernes
    colSabadoInicio = colSabado
    colDomingoInicio = colDomingo
    If colDomingo30 > 0 Then colDomingo30Inicio = colDomingo30
    
    colLunesViernesFin = colSabadoInicio - 1
    colSabadoFin = colDomingoInicio - 1
    
    If colDomingo30Inicio > 0 Then
        colDomingoFin = colDomingo30Inicio - 1
        colDomingo30Fin = colEN - 1   ' Hasta antes de idiomas
    Else
        colDomingoFin = colEN - 1
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
    
    '------------------------------------------------
    ' Última fila de datos (por COD)
    '------------------------------------------------
    ultimaFilaDatos = hojaHorarios.Cells(hojaHorarios.Rows.Count, colCOD).End(xlUp).Row
    
    '------------------------------------------------
    ' Bucle principal por filas de datos
    '------------------------------------------------
    For filaActual = filaPrimeraDatos To ultimaFilaDatos
        
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
        
        ' Escribir en celdas
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
