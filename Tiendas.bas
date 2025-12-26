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

' Día especial (domingos raros, sábados especiales, festivos, etc.)
Type DiaEspecialInfo
    codigo As String           ' Ej: "DOM7", "DOM14", "SAB6"
    cabecera As String         ' Texto cabecera principal en la hoja (ej. "Domingo 7-12")
    cabeceraAlt As String      ' Alternativa (ej. "Domingo 7")
    colInicio As Long
    colFin As Long
    Columnas As ColumnasDiaInfo
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
'   DETECTAR TEXTO "CERRADO"
'====================================================
Private Function EsTextoCerrado(ByVal valor As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim s As String
    s = Trim$(CStr(valor))
    
    ' Acepta "Cerrado" y "CERRADO" (y variantes de mayúsculas)
    EsTextoCerrado = (StrComp(s, "Cerrado", vbTextCompare) = 0)
    Exit Function
ErrHandler:
    EsTextoCerrado = False
End Function

Private Function PalabraCerrado(ByVal idioma As String) As String
    Select Case UCase$(idioma)
        Case "EN": PalabraCerrado = "Closed"
        Case "ES": PalabraCerrado = "Cerrado"
        Case "GL": PalabraCerrado = "Pechado"
        Case "CA": PalabraCerrado = "Tancat"
        Case Else: PalabraCerrado = "Closed"
    End Select
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
    Dim hI1 As String, hF1 As String, hI2 As String, hF2 As String
    
    hI1 = FormatearHora(horario.HoraInicio1)
    hF1 = FormatearHora(horario.HoraFin1)
    hI2 = FormatearHora(horario.HoraInicio2)
    hF2 = FormatearHora(horario.HoraFin2)
    
    ' Primer turno
    If hI1 <> "" And hF1 <> "" Then
        texto = hI1 & " - " & hF1
    End If
    
    ' Segundo turno
    If hI2 <> "" And hF2 <> "" Then
        If texto <> "" Then texto = texto & " / "
        texto = texto & hI2 & " - " & hF2
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
    
    ' Caso 2: L = S ? D
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
            ' Caso 1: L = S = D ? "Lun - Dom: ..."
            ConstruirTextoHorario = prefijoLunesDomingo & ObtenerTextoDia(horarioLunesViernes)
        
        Case 2
            ' Caso 2: L = S ? D
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
'   FORMATO CUANDO L-J Y VIERNES SON DISTINTOS
'====================================================
Private Function ConstruirTextoLunesJuevesViernes( _
    horarioLunesJueves As HorarioDiaInfo, _
    horarioViernes As HorarioDiaInfo, _
    horarioSabado As HorarioDiaInfo, _
    horarioDomingo As HorarioDiaInfo, _
    ByVal idioma As String) As String

    On Error GoTo ErrHandler
    
    Dim prefijoLunesJueves As String
    Dim prefijoViernes As String
    Dim prefijoSabado As String
    Dim prefijoDomingo As String
    Dim prefijoSabDom As String
    Dim separadorBloques As String
    Dim textoResultado As String
    Dim textoParcial As String
    Dim weekendIguales As Boolean
    
    separadorBloques = " | "
    
    ' Prefijos por idioma
    Select Case UCase$(idioma)
        Case "EN"
            prefijoLunesJueves = "Mon - Thu: "
            prefijoViernes = "Fri: "
            prefijoSabado = "Sat: "
            prefijoDomingo = "Sun: "
            prefijoSabDom = "Sat - Sun: "
        Case "ES"
            prefijoLunesJueves = "Lun - Jue: "
            prefijoViernes = "Vie: "
            prefijoSabado = "Sáb: "
            prefijoDomingo = "Dom: "
            prefijoSabDom = "Sáb - Dom: "
        Case "GL"
            prefijoLunesJueves = "Lun - Xov: "
            prefijoViernes = "Ven: "
            prefijoSabado = "Sáb: "
            prefijoDomingo = "Dom: "
            prefijoSabDom = "Sáb - Dom: "
        Case "CA"
            prefijoLunesJueves = "Dil - Dij: "
            prefijoViernes = "Div: "
            prefijoSabado = "Dis: "
            prefijoDomingo = "Diu: "
            prefijoSabDom = "Dis - Diu: "
        Case Else
            prefijoLunesJueves = "Mon - Thu: "
            prefijoViernes = "Fri: "
            prefijoSabado = "Sat: "
            prefijoDomingo = "Sun: "
            prefijoSabDom = "Sat - Sun: "
    End Select
    
    textoResultado = ""
    
    ' Bloque Lunes-Jueves
    If Not EsDiaVacio(horarioLunesJueves) Then
        textoParcial = prefijoLunesJueves & ObtenerTextoDia(horarioLunesJueves)
        If textoParcial <> "" Then
            If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
            textoResultado = textoResultado & textoParcial
        End If
    End If
    
    ' Bloque Viernes
    If Not EsDiaVacio(horarioViernes) Then
        textoParcial = prefijoViernes & ObtenerTextoDia(horarioViernes)
        If textoParcial <> "" Then
            If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
            textoResultado = textoResultado & textoParcial
        End If
    End If
    
    ' ¿Sábado y domingo iguales?
    weekendIguales = (Not EsDiaVacio(horarioSabado) And _
                      Not EsDiaVacio(horarioDomingo) And _
                      SonHorariosIguales(horarioSabado, horarioDomingo))
    
    If weekendIguales Then
        ' Sat - Sun: ...
        textoParcial = prefijoSabDom & ObtenerTextoDia(horarioSabado)
        If textoParcial <> "" Then
            If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
            textoResultado = textoResultado & textoParcial
        End If
    Else
        ' Sábado solo
        If Not EsDiaVacio(horarioSabado) Then
            textoParcial = prefijoSabado & ObtenerTextoDia(horarioSabado)
            If textoParcial <> "" Then
                If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
                textoResultado = textoResultado & textoParcial
            End If
        End If
        
        ' Domingo solo
        If Not EsDiaVacio(horarioDomingo) Then
            textoParcial = prefijoDomingo & ObtenerTextoDia(horarioDomingo)
            If textoParcial <> "" Then
                If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
                textoResultado = textoResultado & textoParcial
            End If
        End If
    End If
    
    ConstruirTextoLunesJuevesViernes = textoResultado
    Exit Function
    
ErrHandler:
    ConstruirTextoLunesJuevesViernes = ""
End Function

'====================================================
'   TEXTO DOMINGO ESPECIAL (DOMINGO 30)
'====================================================
Public Function TextoDiaEspecial(ByVal idioma As String) As String
    On Error GoTo ErrHandler
    
    Select Case UCase$(idioma)
        Case "EN": TextoDiaEspecial = "Sunday Nov 30: "
        Case "ES": TextoDiaEspecial = "Domingo 30 Nov: "
        Case "GL": TextoDiaEspecial = "Domingo 30 Nov: "
        Case "CA": TextoDiaEspecial = "Diumenge 30 Nov: "
        Case Else: TextoDiaEspecial = "Sunday Nov 30: "
    End Select
    
    Exit Function
ErrHandler:
    TextoDiaEspecial = ""
End Function

'====================================================
'   TEXTO DÍAS ESPECIALES (CÓDIGOS SAB6, DOM7, DOM14, DOM21, DOM28...)
'====================================================
Private Function TextoDiaEspecialDic( _
    ByVal idioma As String, _
    ByVal codigoDia As String) As String
    
    On Error GoTo ErrHandler
    
    Select Case UCase$(idioma)
        Case "ES"
            Select Case UCase$(codigoDia)
                Case "DOM21": TextoDiaEspecialDic = "Domingo 21 de Dic: "
                Case "MIE24": TextoDiaEspecialDic = "Miércoles 24 de Dic: "
                Case "VIE26": TextoDiaEspecialDic = "Viernes 26 de Dic: "
                Case "DOM28": TextoDiaEspecialDic = "Domingo 28 de Dic: "
                Case "MIE31": TextoDiaEspecialDic = "Miércoles 31 de Dic: "
            End Select
        
        Case "EN"
            Select Case UCase$(codigoDia)
                Case "DOM21": TextoDiaEspecialDic = "Sunday Dec 21: "
                Case "MIE24": TextoDiaEspecialDic = "Wednesday Dec 24: "
                Case "VIE26": TextoDiaEspecialDic = "Friday Dec 26: "
                Case "DOM28": TextoDiaEspecialDic = "Sunday Dec 28: "
                Case "MIE31": TextoDiaEspecialDic = "Wednesday Dec 31: "
            End Select
        
        Case "GL"
            Select Case UCase$(codigoDia)
                Case "DOM21": TextoDiaEspecialDic = "Domingo 21 de Dec: "
                Case "MIE24": TextoDiaEspecialDic = "Mércores 24 de Dec: "
                Case "VIE26": TextoDiaEspecialDic = "Venres 26 de Dec: "
                Case "DOM28": TextoDiaEspecialDic = "Domingo 28 de Dec: "
                Case "MIE31": TextoDiaEspecialDic = "Mércores 31 de Dec: "
            End Select
        
        Case "CA"
            Select Case UCase$(codigoDia)
                Case "DOM21": TextoDiaEspecialDic = "Diumenge 21 de Des: "
                Case "MIE24": TextoDiaEspecialDic = "Dimecres 24 de Des: "
                Case "VIE26": TextoDiaEspecialDic = "Divendres 26 de Des: "
                Case "DOM28": TextoDiaEspecialDic = "Diumenge 28 de Des: "
                Case "MIE31": TextoDiaEspecialDic = "Dimecres 31 de Des: "
            End Select
        
        Case Else
            Select Case UCase$(codigoDia)
                Case "DOM21": TextoDiaEspecialDic = "Sun Dec 21: "
                Case "MIE24": TextoDiaEspecialDic = "Wed Dec 24: "
                Case "VIE26": TextoDiaEspecialDic = "Fri Dec 26: "
                Case "DOM28": TextoDiaEspecialDic = "Sun Dec 28: "
                Case "MIE31": TextoDiaEspecialDic = "Wed Dec 31: "
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

'====================================================
'   SIGUIENTE INICIO DE BLOQUE (CON DÍAS ESPECIALES)
'====================================================
Private Function SiguienteInicioBloque(ByVal colActual As Long, _
                                       ByVal c1 As Long, ByVal c2 As Long, _
                                       ByVal c3 As Long, ByVal c4 As Long, _
                                       ByVal c5 As Long, ByVal c6 As Long, _
                                       ByRef DiasEspeciales() As DiaEspecialInfo, _
                                       ByVal nEsp As Long) As Long
    On Error GoTo ErrHandler
    
    Dim menor As Long
    Dim c As Long
    Dim i As Long
    
    menor = 0
    
    ' Candidatos "fijos"
    c = c1
    If c <> 0 And c > colActual Then
        If menor = 0 Or c < menor Then menor = c
    End If
    
    c = c2
    If c <> 0 And c > colActual Then
        If menor = 0 Or c < menor Then menor = c
    End If
    
    c = c3
    If c <> 0 And c > colActual Then
        If menor = 0 Or c < menor Then menor = c
    End If
    
    c = c4
    If c <> 0 And c > colActual Then
        If menor = 0 Or c < menor Then menor = c
    End If
    
    c = c5
    If c <> 0 And c > colActual Then
        If menor = 0 Or c < menor Then menor = c
    End If
    
    c = c6
    If c <> 0 And c > colActual Then
        If menor = 0 Or c < menor Then menor = c
    End If
    
    ' Candidatos: días especiales
    If nEsp > 0 Then
        For i = 1 To nEsp
            c = DiasEspeciales(i).colInicio
            If c <> 0 And c > colActual Then
                If menor = 0 Or c < menor Then menor = c
            End If
        Next i
    End If
    
    SiguienteInicioBloque = menor
    Exit Function
ErrHandler:
    SiguienteInicioBloque = 0
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
    columnasDia As ColumnasDiaInfo, _
    Optional ByVal permitirTextoCerrado As Boolean = False) As HorarioDiaInfo
    
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
    
    ' Caso especial: texto "Cerrado" en la primera apertura (solo si se permite)
    If permitirTextoCerrado Then
        If EsTextoCerrado(apertura1) Then
            horario.HoraInicio1 = "Cerrado"
            horario.HoraFin1 = Empty
            horario.HoraInicio2 = Empty
            horario.HoraFin2 = Empty
            LeerHorarioDeDia = horario
            Exit Function
        End If
    End If
    
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
        If FormatearHora(apertura1) <> "" Then
            horaInicioUnica = apertura1
        End If
        
        If FormatearHora(apertura2) <> "" Then
            If IsEmpty(horaInicioUnica) Or apertura2 < horaInicioUnica Then
                horaInicioUnica = apertura2
            End If
        End If
        
        ' Cierre más tardío
        If FormatearHora(cierre1) <> "" Then
            horaFinUnica = cierre1
        End If
        
        If FormatearHora(cierre2) <> "" Then
            If IsEmpty(horaFinUnica) Or cierre2 > horaFinUnica Then
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
'   AÑADIR DÍA ESPECIAL AL ARRAY
'====================================================
Private Sub AddDiaEspecial(ByRef DiasEspeciales() As DiaEspecialInfo, _
                           ByRef nEsp As Long, _
                           ByVal codigo As String, _
                           ByVal cabecera As String, _
                           Optional ByVal cabeceraAlt As String = "")
    nEsp = nEsp + 1
    ReDim Preserve DiasEspeciales(1 To nEsp)
    
    With DiasEspeciales(nEsp)
        .codigo = codigo
        .cabecera = cabecera
        .cabeceraAlt = cabeceraAlt
    End With
End Sub

'====================================================
'   TEXTO DE TODOS LOS DÍAS ESPECIALES PARA UNA FILA
'====================================================
Private Function TextoDiasEspeciales(horEsp() As HorarioDiaInfo, _
                                     DiasEspeciales() As DiaEspecialInfo, _
                                     ByVal idioma As String) As String
    Dim i As Long
    Dim txt As String
    Dim bloque As String
    
    For i = LBound(horEsp) To UBound(horEsp)
        
        ' Si el día especial está marcado como "Cerrado"
        If EsTextoCerrado(horEsp(i).HoraInicio1) Then
            bloque = TextoDiaEspecialDic(idioma, DiasEspeciales(i).codigo) & PalabraCerrado(idioma)
        
        ElseIf Not EsDiaVacio(horEsp(i)) Then
            bloque = TextoDiaEspecialDic(idioma, DiasEspeciales(i).codigo) & ObtenerTextoDia(horEsp(i))
        Else
            bloque = ""
        End If
        
        If bloque <> "" Then
            If txt <> "" Then txt = txt & " | "
            txt = txt & bloque
        End If
    Next i
    
    TextoDiasEspeciales = txt
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
    Dim i As Long
    
    '========================================
    '   Estructuras de horarios
    '========================================
    Dim horarioLunesJueves As HorarioDiaInfo
    Dim horarioLunesViernes As HorarioDiaInfo
    Dim horarioViernes As HorarioDiaInfo
    Dim horarioSabado As HorarioDiaInfo
    Dim horarioDomingo As HorarioDiaInfo
    
    Dim casoHorario As Integer
    
    '========================================
    '   Configuración de filas
    '========================================
    Dim filaCabeceraPrincipal As Long
    Dim filaSubcabeceraInicio As Long
    Dim filaSubcabeceraFin As Long
    Dim filaPrimeraDatos As Long
    
    '========================================
    '   Columnas cabecera de días
    '========================================
    Dim colCOD As Long
    Dim colLunesJueves As Long
    Dim colLunesViernes As Long
    Dim colViernes As Long
    Dim colSabado As Long
    Dim colDomingo As Long
    
    ' Columnas idiomas
    Dim colEN As Long, colES As Long, colGL As Long, colCA As Long
    
    ' Rangos de bloques de columnas
    Dim colLunesJuevesInicio As Long, colLunesJuevesFin As Long
    Dim colLunesViernesInicio As Long, colLunesViernesFin As Long
    Dim colViernesInicio As Long, colViernesFin As Long
    Dim colSabadoInicio As Long, colSabadoFin As Long
    Dim colDomingoInicio As Long, colDomingoFin As Long
    
    ' Info de columnas de apertura/cierre por bloque
    Dim columnasLunesJueves As ColumnasDiaInfo
    Dim columnasLunesViernes As ColumnasDiaInfo
    Dim columnasViernes As ColumnasDiaInfo
    Dim columnasSabado As ColumnasDiaInfo
    Dim columnasDomingo As ColumnasDiaInfo
    
    ' Días especiales en array
    Dim DiasEspeciales() As DiaEspecialInfo
    Dim nEsp As Long
    
    ' Textos resultado
    Dim textoEN As String, textoES As String, textoGL As String, textoCA As String
    Dim celdaCabeceraCOD As Range
    Dim celdaTmp As Range
    
    ' Acumuladores días especiales
    Dim extrasEN As String, extrasES As String, extrasGL As String, extrasCA As String
    
    ' Flags para saber qué formato usamos
    Dim usaLunesJuevesMasViernes As Boolean
    Dim usaLunesViernes As Boolean
    
    Dim viernesVacio As Boolean
    Dim viernesIgualLJ As Boolean
    
    ' Array de horarios especiales por fila
    Dim horEsp() As HorarioDiaInfo
    
    '------------------------------------------------
    ' Localizar hoja
    '------------------------------------------------
    On Error Resume Next

        Set hojaHorarios = ThisWorkbook.Worksheets("HORARIO INTERNACIONAL")
        
    On Error GoTo ErrHandler
    
    If hojaHorarios Is Nothing Then
        MsgBox "No se ha encontrado la hoja de horarios (""Horarios habituales"", ""HORARIO ESPAÑA"" o ""HORARIO INTERNACIONAL"").", _
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
        MsgBox "No se han encontrado correctamente todas las columnas de idioma (Inglés/Español).", _
               vbCritical, "Horarios"
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Bloques de días (cabeceras normales)
    '------------------------------------------------
    colLunesJueves = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Lunes a Jueves")
    colViernes = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Viernes")
    colLunesViernes = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Lunes a Viernes")
    colSabado = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Sábado")
    colDomingo = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Domingo")
    
    '------------------------------------------------
    ' Definir días especiales (DOMINGOS RAROS, etc.)
    '   -> Aquí añadir/quitar días especiales
    '------------------------------------------------
    nEsp = 0
    ReDim DiasEspeciales(1 To 1)
    
    ' Ejemplos diciembre
    AddDiaEspecial DiasEspeciales, nEsp, "DOM21", "Domingo 21-12", "Domingo 21"
    AddDiaEspecial DiasEspeciales, nEsp, "MIE24", "Miércoles 24-12", "Miércoles 24"
    AddDiaEspecial DiasEspeciales, nEsp, "VIE26", "Viernes 26-12", "Viernes 26"
    AddDiaEspecial DiasEspeciales, nEsp, "DOM28", "Domingo 28-12", "Domingo 28"
    AddDiaEspecial DiasEspeciales, nEsp, "MIE31", "Miércoles 31-12", "Miércoles 31"
    
    ' Localizar columnas de cada día especial
    If nEsp > 0 Then
        For i = 1 To nEsp
            With DiasEspeciales(i)
                .colInicio = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, .cabecera)
                If .colInicio = 0 And .cabeceraAlt <> "" Then
                    .colInicio = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, .cabeceraAlt)
                End If
            End With
        Next i
    End If
    
    '------------------------------------------------
    ' ¿Qué formato usamos?
    '------------------------------------------------
    usaLunesJuevesMasViernes = (colLunesJueves <> 0 And colViernes <> 0)
    usaLunesViernes = (colLunesViernes <> 0)
    
    If Not usaLunesJuevesMasViernes And Not usaLunesViernes Then
        If colLunesJueves <> 0 Then
            ' Caso especial: hay "Lunes a Jueves" pero no "Viernes".
            ' Suponemos que el viernes tiene el mismo horario que L-J
            usaLunesViernes = True
            colLunesViernes = colLunesJueves
        Else
            MsgBox "No se han encontrado correctamente las columnas de lunes (""Lunes a Jueves""/""Viernes"" o ""Lunes a Viernes"").", _
                   vbCritical, "Horarios"
            Exit Sub
        End If
    End If
    
    If colSabado = 0 Then
        MsgBox "No se ha encontrado correctamente la columna de ""Sábado"".", _
               vbCritical, "Horarios"
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Rangos de columnas para cada bloque
    '------------------------------------------------
    If usaLunesJuevesMasViernes Then
        ' Formato España
        colLunesJuevesInicio = colLunesJueves
        colLunesJuevesFin = SiguienteInicioBloque(colLunesJuevesInicio, _
                                                  colViernes, colSabado, colDomingo, colEN, 0, 0, _
                                                  DiasEspeciales, nEsp) - 1
        
        colViernesInicio = colViernes
        colViernesFin = SiguienteInicioBloque(colViernesInicio, _
                                              colSabado, colDomingo, colEN, 0, 0, 0, _
                                              DiasEspeciales, nEsp) - 1
    ElseIf usaLunesViernes Then
        ' Formato Internacional
        colLunesViernesInicio = colLunesViernes
        colLunesViernesFin = SiguienteInicioBloque(colLunesViernesInicio, _
                                                   colSabado, colDomingo, colEN, 0, 0, 0, _
                                                   DiasEspeciales, nEsp) - 1
    End If
    
    colSabadoInicio = colSabado
    colSabadoFin = SiguienteInicioBloque(colSabadoInicio, _
                                         colDomingo, colEN, 0, 0, 0, 0, _
                                         DiasEspeciales, nEsp) - 1
    
    If colDomingo > 0 Then
        colDomingoInicio = colDomingo
        colDomingoFin = SiguienteInicioBloque(colDomingoInicio, _
                                              colEN, 0, 0, 0, 0, 0, _
                                              DiasEspeciales, nEsp) - 1
    Else
        colDomingoInicio = 0
        colDomingoFin = 0
    End If
    
    ' Rangos para cada día especial
    If nEsp > 0 Then
        For i = 1 To nEsp
            If DiasEspeciales(i).colInicio > 0 Then
                DiasEspeciales(i).colFin = SiguienteInicioBloque(DiasEspeciales(i).colInicio, _
                                                                 colEN, 0, 0, 0, 0, 0, _
                                                                 DiasEspeciales, nEsp) - 1
            End If
        Next i
    End If
    
    '------------------------------------------------
    ' Detectar columnas de Apertura/Cierre por bloque
    '------------------------------------------------
    If usaLunesJuevesMasViernes Then
        columnasLunesJueves = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                                colLunesJuevesInicio, colLunesJuevesFin)
        columnasViernes = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                            colViernesInicio, colViernesFin)
    ElseIf usaLunesViernes Then
        columnasLunesViernes = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                                 colLunesViernesInicio, colLunesViernesFin)
    End If
    
    columnasSabado = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                       colSabadoInicio, colSabadoFin)
    columnasDomingo = DetectarColumnasAperturaCierreDia(hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                                                        colDomingoInicio, colDomingoFin)
    
    ' Detectar columnas para todos los días especiales
    If nEsp > 0 Then
        For i = 1 To nEsp
            If DiasEspeciales(i).colInicio > 0 And _
               DiasEspeciales(i).colFin >= DiasEspeciales(i).colInicio Then
                DiasEspeciales(i).Columnas = DetectarColumnasAperturaCierreDia( _
                    hojaHorarios, filaSubcabeceraInicio, filaSubcabeceraFin, _
                    DiasEspeciales(i).colInicio, DiasEspeciales(i).colFin)
            End If
        Next i
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
        
        ' Leer horarios por día, según formato
        If usaLunesJuevesMasViernes Then
            horarioLunesJueves = LeerHorarioDeDia(hojaHorarios, filaActual, columnasLunesJueves)
            horarioViernes = LeerHorarioDeDia(hojaHorarios, filaActual, columnasViernes)
        ElseIf usaLunesViernes Then
            horarioLunesViernes = LeerHorarioDeDia(hojaHorarios, filaActual, columnasLunesViernes)
        End If
        
        horarioSabado = LeerHorarioDeDia(hojaHorarios, filaActual, columnasSabado)
        horarioDomingo = LeerHorarioDeDia(hojaHorarios, filaActual, columnasDomingo)
        
        '--------------------------------------------
        ' Construcción del texto principal (EN/ES/GL/CA)
        '--------------------------------------------
        If usaLunesJuevesMasViernes Then
            ' Formato España: "Lun - Jue" y "Vie"
            viernesVacio = (columnasViernes.NumTurnos = 0 Or EsDiaVacio(horarioViernes))
            viernesIgualLJ = (Not viernesVacio And SonHorariosIguales(horarioLunesJueves, horarioViernes))
            
            If viernesVacio Or viernesIgualLJ Then
                ' Viernes igual que Lunes-Jueves -> tratamos como Lunes-Viernes
                horarioLunesViernes = horarioLunesJueves
                casoHorario = ObtenerCasoHorario(horarioLunesViernes, horarioSabado, horarioDomingo)
                
                textoEN = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "EN")
                textoES = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "ES")
                textoGL = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "GL")
                textoCA = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "CA")
            Else
                ' Viernes diferente -> "Mon - Thu: ..." | "Fri: ..."
                textoEN = ConstruirTextoLunesJuevesViernes(horarioLunesJueves, horarioViernes, horarioSabado, horarioDomingo, "EN")
                textoES = ConstruirTextoLunesJuevesViernes(horarioLunesJueves, horarioViernes, horarioSabado, horarioDomingo, "ES")
                textoGL = ConstruirTextoLunesJuevesViernes(horarioLunesJueves, horarioViernes, horarioSabado, horarioDomingo, "GL")
                textoCA = ConstruirTextoLunesJuevesViernes(horarioLunesJueves, horarioViernes, horarioSabado, horarioDomingo, "CA")
            End If
        
        ElseIf usaLunesViernes Then
            ' Formato Internacional: "Lun - Vie"
            casoHorario = ObtenerCasoHorario(horarioLunesViernes, horarioSabado, horarioDomingo)
            
            textoEN = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "EN")
            textoES = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "ES")
            textoGL = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "GL")
            textoCA = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "CA")
        End If
        
        '------------------------------------------------
        ' DÍAS ESPECIALES (ARRAY)
        '------------------------------------------------
        If nEsp > 0 Then
            ReDim horEsp(1 To nEsp)
            
            For i = 1 To nEsp
                If DiasEspeciales(i).colInicio > 0 And DiasEspeciales(i).Columnas.NumTurnos > 0 Then
                    horEsp(i) = LeerHorarioDeDia(hojaHorarios, filaActual, DiasEspeciales(i).Columnas, True)
                Else
                    ' Inicializar vacío
                    With horEsp(i)
                        .HoraInicio1 = Empty: .HoraFin1 = Empty
                        .HoraInicio2 = Empty: .HoraFin2 = Empty
                    End With
                End If
            Next i
            
            extrasEN = TextoDiasEspeciales(horEsp, DiasEspeciales, "EN")
            extrasES = TextoDiasEspeciales(horEsp, DiasEspeciales, "ES")
            extrasGL = TextoDiasEspeciales(horEsp, DiasEspeciales, "GL")
            extrasCA = TextoDiasEspeciales(horEsp, DiasEspeciales, "CA")
        End If
        
        ' Añadir especiales en nueva línea
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
