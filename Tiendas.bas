Option Explicit

' ============================
'  ESTRUCTURAS DE DATOS
' ============================
' HorarioDiaInfo:
'   Guarda hasta 2 turnos (por ejemplo mañana/tarde) para un mismo "día/bloque".
'   
'   Se usa Variant porque en Excel puedes tener:
'     - horas reales (Date)
'     - vacíos
'     - texto (p.ej. "Cerrado" en días especiales)
Type HorarioDiaInfo
    HoraInicio1 As Variant
    HoraFin1 As Variant
    HoraInicio2 As Variant
    HoraFin2 As Variant
End Type

' ColumnasDiaInfo:
'   Describe dónde están las columnas de "Apertura" y "Cierre" dentro de un bloque
'   (Lunes a Viernes, Sábado, etc.). Soporta hasta 2 turnos.
Type ColumnasDiaInfo
    ColApertura(1 To 2) As Long
    ColCierre(1 To 2) As Long
    NumTurnos As Integer
End Type

' DiaEspecialInfo:
'   Metadatos para bloques especiales (31 dic, 4 ene, ...).
'   - codigo: identificador interno para sacar el texto traducido
'   - cabecera/cabeceraAlt: nombres posibles en la cabecera del Excel
'   - colInicio/colFin: rango de columnas del bloque en la hoja
'   - Columnas: subrango de columnas Apertura/Cierre detectadas
Type DiaEspecialInfo
    codigo As String
    cabecera As String
    cabeceraAlt As String
    colInicio As Long
    colFin As Long
    Columnas As ColumnasDiaInfo
End Type

' ============================
'  UTILIDADES GENERALES
' ============================

' Manejo de error homogéneo (para no tener MsgBox diferentes por todos lados)
Private Sub GestionarError(ByVal nombreProcedimiento As String)
    MsgBox "Se ha producido un error en: " & nombreProcedimiento & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en macro de horarios"
End Sub

' Convierte una celda a "hh:mm".
' Si viene vacío, error o algo no formateable, devuelve "" para simplificar comparaciones.
Private Function FormatearHora(ByVal valor As Variant) As String
    On Error GoTo ErrHandler
    
    If IsError(valor) Or IsEmpty(valor) Or valor = "" Then
        FormatearHora = ""
    Else
        FormatearHora = Format$(valor, "hh:mm")
    End If
    
    Exit Function
ErrHandler:
    FormatearHora = ""
End Function

' Devuelve True si una celda contiene algo parecido a "cerrado".
' Lo hacemos por "cerrad" para cubrir Cerrado/Cerrada/Cerrados, etc.
Private Function EsTextoCerrado(ByVal valor As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim s As String
    s = LCase$(Trim$(CStr(valor)))

    EsTextoCerrado = (InStr(1, s, "cerrad", vbTextCompare) > 0)
    Exit Function
ErrHandler:
    EsTextoCerrado = False
End Function

' Traducción de la palabra "Cerrado" según idioma.
Private Function PalabraCerrado(ByVal idioma As String) As String
    Select Case UCase$(idioma)
        Case "EN": PalabraCerrado = "Closed"
        Case "ES": PalabraCerrado = "Cerrado"
        Case "GL": PalabraCerrado = "Pechado"
        Case "CA": PalabraCerrado = "Tancat"
        Case Else: PalabraCerrado = "Closed"
    End Select
End Function

' Compara horarios normalizando horas a "hh:mm".
' (Así da igual si Excel guarda la hora como Date y otro viene como texto con formato)
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

' Día/bloque "vacío" = no hay horas en ninguno de los 2 turnos.
Private Function EsDiaVacio(horario As HorarioDiaInfo) As Boolean
    On Error GoTo ErrHandler
    
    EsDiaVacio = (FormatearHora(horario.HoraInicio1) = "" And _
                  FormatearHora(horario.HoraFin1) = "" And _
                  FormatearHora(horario.HoraInicio2) = "" And _
                  FormatearHora(horario.HoraFin2) = "")
    Exit Function
ErrHandler:
    ' Si falla el formateo por lo que sea, consideramos vacío para no reventar el texto.
    EsDiaVacio = True
End Function

' Construye el texto "08:00 - 14:00 / 16:00 - 20:00" para un HorarioDiaInfo.
' Si solo hay un turno, devuelve solo ese. Si no hay horas, devuelve "".
Private Function ObtenerTextoDia(horario As HorarioDiaInfo) As String
    On Error GoTo ErrHandler
    
    Dim texto As String
    Dim hI1 As String, hF1 As String, hI2 As String, hF2 As String
    
    hI1 = FormatearHora(horario.HoraInicio1)
    hF1 = FormatearHora(horario.HoraFin1)
    hI2 = FormatearHora(horario.HoraInicio2)
    hF2 = FormatearHora(horario.HoraFin2)
    
    If hI1 <> "" And hF1 <> "" Then
        texto = hI1 & " - " & hF1
    End If
    
    If hI2 <> "" And hF2 <> "" Then
        If texto <> "" Then texto = texto & " / "
        texto = texto & hI2 & " - " & hF2
    End If
    
    ObtenerTextoDia = texto
    Exit Function
ErrHandler:
    ObtenerTextoDia = ""
End Function

' Clasifica el "patrón" de horarios:
'   1 = L-V = Sáb = Dom  -> se puede mostrar "Lun-Dom: ..."
'   2 = L-V = Sáb, Dom distinto -> se puede mostrar "Lun-Sáb: ... | Dom: ..."
'   0 = resto (cada uno por separado si aplica)
Private Function ObtenerCasoHorario( _
    horarioLunesViernes As HorarioDiaInfo, _
    horarioSabado As HorarioDiaInfo, _
    horarioDomingo As HorarioDiaInfo) As Integer
    
    On Error GoTo ErrHandler
    
    If SonHorariosIguales(horarioLunesViernes, horarioSabado) And _
       SonHorariosIguales(horarioLunesViernes, horarioDomingo) Then
        ObtenerCasoHorario = 1
        Exit Function
    End If
    
    If SonHorariosIguales(horarioLunesViernes, horarioSabado) And _
       Not SonHorariosIguales(horarioLunesViernes, horarioDomingo) Then
        ObtenerCasoHorario = 2
        Exit Function
    End If
    
    ObtenerCasoHorario = 0
    Exit Function
ErrHandler:
    ObtenerCasoHorario = 0
End Function

' Construye el texto final para el caso "Lun-Vie" (o "Lun-Sáb" / "Lun-Dom") + fin de semana.
' El objetivo es generar una frase corta cuando hay patrones repetidos.
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
    
    ' Prefijos por idioma (solo texto, la lógica es la misma)
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
            ' Todo igual: un único bloque Lun-Dom
            ConstruirTextoHorario = prefijoLunesDomingo & ObtenerTextoDia(horarioLunesViernes)
        
        Case 2
            ' Lun-Sáb igual, domingo distinto (o vacío)
            If EsDiaVacio(horarioDomingo) Then
                ConstruirTextoHorario = prefijoLunesSabado & ObtenerTextoDia(horarioLunesViernes)
            Else
                ConstruirTextoHorario = prefijoLunesSabado & ObtenerTextoDia(horarioLunesViernes) & _
                                        separadorBloques & prefijoDomingo & ObtenerTextoDia(horarioDomingo)
            End If
        
        Case Else
            ' Caso general: se van añadiendo los bloques que tengan contenido
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

' Variante del constructor cuando el Excel viene separado en:
'   - Lunes a Jueves
'   - Viernes
'   - Sábado
'   - Domingo
' Con compactación del fin de semana si Sáb=Dom.
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
    
    ' Añadir L-J si hay contenido
    If Not EsDiaVacio(horarioLunesJueves) Then
        textoParcial = prefijoLunesJueves & ObtenerTextoDia(horarioLunesJueves)
        If textoParcial <> "" Then
            If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
            textoResultado = textoResultado & textoParcial
        End If
    End If
    
    ' Añadir Viernes si hay contenido
    If Not EsDiaVacio(horarioViernes) Then
        textoParcial = prefijoViernes & ObtenerTextoDia(horarioViernes)
        If textoParcial <> "" Then
            If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
            textoResultado = textoResultado & textoParcial
        End If
    End If
    
    ' Compactar Sáb-Dom si son iguales y ambos existen
    weekendIguales = (Not EsDiaVacio(horarioSabado) And _
                      Not EsDiaVacio(horarioDomingo) And _
                      SonHorariosIguales(horarioSabado, horarioDomingo))
    
    If weekendIguales Then
        textoParcial = prefijoSabDom & ObtenerTextoDia(horarioSabado)
        If textoParcial <> "" Then
            If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
            textoResultado = textoResultado & textoParcial
        End If
    Else
        ' Si no son iguales, se muestran por separado (si aplica)
        If Not EsDiaVacio(horarioSabado) Then
            textoParcial = prefijoSabado & ObtenerTextoDia(horarioSabado)
            If textoParcial <> "" Then
                If textoResultado <> "" Then textoResultado = textoResultado & separadorBloques
                textoResultado = textoResultado & textoParcial
            End If
        End If
        
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

' Traduce el "título" de cada día especial según idioma y código.
' Ojo: aquí los códigos están acoplados a AddDiaEspecial del procedimiento principal.
Private Function TextoDiaEspecial( _
    ByVal idioma As String, _
    ByVal codigoDia As String) As String
    
    On Error GoTo ErrHandler
    
    Select Case UCase$(idioma)
        Case "ES"
            Select Case UCase$(codigoDia)
                Case "MIE31": TextoDiaEspecial = "Miércoles 31 de Dic: "
                Case "DOM04": TextoDiaEspecial = "Domingo 4 de Ene: "
                Case "LUN05": TextoDiaEspecial = "Lunes 5 de Ene: "
                Case "DOM11": TextoDiaEspecial = "Domingo 11 de Ene: "
            End Select
        
        Case "EN"
            Select Case UCase$(codigoDia)
                Case "MIE31": TextoDiaEspecial = "Wednesday Dec 31: "
                Case "DOM04": TextoDiaEspecial = "Sunday Jan 4: "
                Case "LUN05": TextoDiaEspecial = "Monday Jan 5: "
                Case "DOM11": TextoDiaEspecial = "Sunday Jan 11: "
            End Select
        
        Case "GL"
            Select Case UCase$(codigoDia)
                Case "MIE31": TextoDiaEspecial = "Mércores 31 de Dec: "
                Case "DOM04": TextoDiaEspecial = "Domingo 4 de Xan: "
                Case "LUN05": TextoDiaEspecial = "Luns 5 de Xan: "
                Case "DOM11": TextoDiaEspecial = "Domingo 11 de Xan: "
            End Select
        
        Case "CA"
            Select Case UCase$(codigoDia)
                Case "MIE31": TextoDiaEspecial = "Dimecres 31 de Des: "
                Case "DOM04": TextoDiaEspecial = "Diumenge 4 de Gen: "
                Case "LUN05": TextoDiaEspecial = "Dilluns 5 de Gen: "
                Case "DOM11": TextoDiaEspecial = "Diumenge 11 de Gen: "
            End Select
        
        Case Else
            Select Case UCase$(codigoDia)
                Case "MIE31": TextoDiaEspecial = "Wednesday Dec 31: "
                Case "DOM04": TextoDiaEspecial = "Sunday Jan 4: "
                Case "LUN05": TextoDiaEspecial = "Monday Jan 5: "
                Case "DOM11": TextoDiaEspecial = "Sunday Jan 11: "
            End Select
    End Select
    
    Exit Function
ErrHandler:
    TextoDiaEspecial = ""
End Function

' Busca una columna en una fila concreta que sea EXACTAMENTE igual al título indicado.
' Devuelve 0 si no la encuentra.
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

' Busca un texto dentro de un bloque de filas (cabecera y subcabeceras).
' Útil cuando la tabla tiene varias filas de título.
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

' Dado un inicio de bloque (colActual) y varios posibles "siguientes inicios",
' devuelve la columna más pequeña que sea > colActual.
' Se usa para calcular el final del bloque como (siguienteInicio - 1).
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
    
    ' Lista fija (bloques estándar e idiomas)
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
    
    ' Y además tenemos los bloques de días especiales (si existen)
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

' Detecta, dentro de un rango de columnas (colInicio..colFin), cuáles son "Apertura" y "Cierre".
' Busca el texto en la(s) fila(s) de subcabecera.
' - Soporta 1 o 2 turnos.
' - Al final recalcula NumTurnos como número de pares completos (Apertura+Cierre).
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
                If InStr(1, textoCelda, "apert", vbTextCompare) > 0 Then
                    ' Guardamos hasta 2 aperturas (turno 1 y turno 2)
                    If infoColumnas.NumTurnos < 2 Then
                        infoColumnas.NumTurnos = infoColumnas.NumTurnos + 1
                        infoColumnas.ColApertura(infoColumnas.NumTurnos) = col
                    End If
                    Exit For
                ElseIf InStr(1, textoCelda, "cierre", vbTextCompare) > 0 Then
                    ' Guardamos el cierre en el primer hueco disponible (máx 2)
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
    
    ' Recalcular turnos como pares completos encontrados
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
    DetectarColumnasAperturaCierreDia = infoColumnas
End Function

' Lee las horas de una fila para el "día/bloque" descrito por columnasDia.
' - Si hay dos turnos completos, se respetan como Turno1 y Turno2.
' - Si vienen datos incompletos/repartidos (muy típico en Excels), intenta:
'     inicio = el menor de las aperturas válidas
'     fin    = el mayor de los cierres válidos
' - permitirTextoCerrado: en días especiales se admite que pongan "Cerrado" en una celda.
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
    
    ' Si se permite, interpretamos cualquier "cerrado" como día cerrado completo
    If permitirTextoCerrado Then
        If EsTextoCerrado(apertura1) Or EsTextoCerrado(cierre1) Or _
           EsTextoCerrado(apertura2) Or EsTextoCerrado(cierre2) Then
            horario.HoraInicio1 = "Cerrado"
            horario.HoraFin1 = Empty
            horario.HoraInicio2 = Empty
            horario.HoraFin2 = Empty
            LeerHorarioDeDia = horario
            Exit Function
        End If
    End If

    ' Turno 2 (si existe)
    If columnasDia.NumTurnos >= 2 Then
        If columnasDia.ColApertura(2) <> 0 Then apertura2 = hoja.Cells(fila, columnasDia.ColApertura(2)).Value
        If columnasDia.ColCierre(2) <> 0 Then cierre2 = hoja.Cells(fila, columnasDia.ColCierre(2)).Value
    End If
    
    ' Caso "limpio": dos turnos completos
    If FormatearHora(apertura1) <> "" And FormatearHora(cierre1) <> "" And _
       FormatearHora(apertura2) <> "" And FormatearHora(cierre2) <> "" Then
       
        horaInicioUnica = apertura1: horaFinUnica = cierre1
        horaInicio2Turno = apertura2: horaFin2Turno = cierre2
    
    Else
        ' Caso "sucio": consolidar a un único tramo usando min(aperturas) y max(cierres)
        If FormatearHora(apertura1) <> "" Then
            horaInicioUnica = apertura1
        End If
        
        If FormatearHora(apertura2) <> "" Then
            If IsEmpty(horaInicioUnica) Or apertura2 < horaInicioUnica Then
                horaInicioUnica = apertura2
            End If
        End If
        
        If FormatearHora(cierre1) <> "" Then
            horaFinUnica = cierre1
        End If
        
        If FormatearHora(cierre2) <> "" Then
            If IsEmpty(horaFinUnica) Or cierre2 > horaFinUnica Then
                horaFinUnica = cierre2
            End If
        End If
    End If
    
    horario.HoraInicio1 = horaInicioUnica
    horario.HoraFin1 = horaFinUnica
    horario.HoraInicio2 = horaInicio2Turno
    horario.HoraFin2 = horaFin2Turno
    
    LeerHorarioDeDia = horario
    Exit Function
ErrHandler:
    LeerHorarioDeDia = horario
End Function

' Helper para ir metiendo días especiales en un array dinámico.
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

' Construye el texto de TODOS los días especiales encontrados para una fila:
'   "Wednesday Dec 31: Closed | Sunday Jan 4: 09:00 - 14:00"
Private Function TextoDiasEspeciales(horEsp() As HorarioDiaInfo, _
                                     DiasEspeciales() As DiaEspecialInfo, _
                                     ByVal idioma As String) As String
    Dim i As Long
    Dim txt As String
    Dim bloque As String
    
    For i = LBound(horEsp) To UBound(horEsp)
        If EsTextoCerrado(horEsp(i).HoraInicio1) Then
            bloque = TextoDiaEspecial(idioma, DiasEspeciales(i).codigo) & PalabraCerrado(idioma)
        ElseIf Not EsDiaVacio(horEsp(i)) Then
            bloque = TextoDiaEspecial(idioma, DiasEspeciales(i).codigo) & ObtenerTextoDia(horEsp(i))
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

' ============================
'  PROCEDIMIENTO PRINCIPAL
' ============================
' Recorre la hoja de horarios, detecta dónde están los bloques (L-V / L-J+V / Sáb / Dom),
' lee las horas de cada fila, y escribe el texto final en columnas de idioma.
Public Sub Horarios()

    Const NOMBRE_PROC As String = "Horarios"
    On Error GoTo ErrHandler
    
    Dim hojaHorarios As Worksheet
    Dim ultimaFilaDatos As Long
    Dim filaActual As Long
    Dim i As Long
    
    ' ============================
    '  ESTRUCTURAS DE HORARIOS
    ' ============================
    Dim horarioLunesJueves As HorarioDiaInfo
    Dim horarioLunesViernes As HorarioDiaInfo
    Dim horarioViernes As HorarioDiaInfo
    Dim horarioSabado As HorarioDiaInfo
    Dim horarioDomingo As HorarioDiaInfo
    
    Dim casoHorario As Integer
    
    ' ============================
    '  FILAS DE CABECERA
    ' ============================
    ' La tabla tiene varias filas de título:
    '   - Fila principal con "COD", "TIENDA", "Horario Habitual", etc.
    '   - Filas intermedias con bloques ("Lunes a Viernes", "Sábado", "Domingo", días especiales)
    '   - Subcabeceras con "Mañana / Tarde" y "Apertura / Cierre"
    Dim filaCabeceraPrincipal As Long
    Dim filaSubcabeceraInicio As Long
    Dim filaSubcabeceraFin As Long
    Dim filaPrimeraDatos As Long
    
    ' ============================
    '  COLUMNAS PRINCIPALES
    ' ============================
    Dim colCOD As Long
    Dim colLunesJueves As Long
    Dim colLunesViernes As Long
    Dim colViernes As Long
    Dim colSabado As Long
    Dim colDomingo As Long
    
    ' Columnas destino de texto final (idiomas)
    ' En este Excel están al final: Inglés | Catalán | Gallego | Español
    Dim colEN As Long, colES As Long, colGL As Long, colCA As Long
    
    ' ============================
    '  RANGOS DE BLOQUES (inicio/fin)
    ' ============================
    Dim colLunesJuevesInicio As Long, colLunesJuevesFin As Long
    Dim colLunesViernesInicio As Long, colLunesViernesFin As Long
    Dim colViernesInicio As Long, colViernesFin As Long
    Dim colSabadoInicio As Long, colSabadoFin As Long
    Dim colDomingoInicio As Long, colDomingoFin As Long
    
    ' Columnas de Apertura/Cierre detectadas dentro de cada bloque
    Dim columnasLunesJueves As ColumnasDiaInfo
    Dim columnasLunesViernes As ColumnasDiaInfo
    Dim columnasViernes As ColumnasDiaInfo
    Dim columnasSabado As ColumnasDiaInfo
    Dim columnasDomingo As ColumnasDiaInfo
    
    ' ============================
    '  DÍAS ESPECIALES
    ' ============================
    ' En esta plantilla los días especiales son bloques completos
    ' con el mismo formato que el horario habitual (Mañana/Tarde + Apertura/Cierre)
    Dim DiasEspeciales() As DiaEspecialInfo
    Dim nEsp As Long
    
    ' ============================
    '  TEXTOS FINALES
    ' ============================
    Dim textoEN As String, textoES As String, textoGL As String, textoCA As String
    Dim celdaCabeceraCOD As Range
    Dim celdaTmp As Range
    
    ' Texto adicional para días especiales (se añade en una línea aparte)
    Dim extrasEN As String, extrasES As String, extrasGL As String, extrasCA As String
    
    ' Flags para detectar el formato de la hoja:
    '   - Lunes a Viernes (formato habitual en esta plantilla)
    '   - Lunes a Jueves + Viernes (soportado pero menos común aquí)
    Dim usaLunesJuevesMasViernes As Boolean
    Dim usaLunesViernes As Boolean
    
    Dim viernesVacio As Boolean
    Dim viernesIgualLJ As Boolean
    
    Dim horEsp() As HorarioDiaInfo
    
    ' ============================
    '  LOCALIZAR HOJA
    ' ============================
    On Error Resume Next
        Set hojaHorarios = ThisWorkbook.Worksheets("HORARIO ESPAÑA")
    On Error GoTo ErrHandler
    
    If hojaHorarios Is Nothing Then
        MsgBox "No se ha encontrado la hoja de horarios (""HORARIO ESPAÑA"").", _
               vbCritical, "Horarios"
        Exit Sub
    End If
    
    ' ============================
    '  LOCALIZAR CABECERA "COD"
    ' ============================
    ' "COD" marca el inicio lógico de la tabla.
    ' A partir de aquí se detectan:
    '   - filas de subcabecera
    '   - primera fila de datos real
    Set celdaCabeceraCOD = hojaHorarios.Cells.Find(What:="COD", LookIn:=xlValues, LookAt:=xlWhole, _
                                                   SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If celdaCabeceraCOD Is Nothing Then
        MsgBox "No se ha encontrado la cabecera 'COD' en la hoja de horarios.", vbCritical, "Horarios"
        Exit Sub
    End If
    
    filaCabeceraPrincipal = celdaCabeceraCOD.Row
    colCOD = celdaCabeceraCOD.Column
    
    filaSubcabeceraInicio = filaCabeceraPrincipal + 1
    
    ' Primera fila con datos = primera fila donde COD ya tiene valor
    filaPrimeraDatos = filaCabeceraPrincipal + 1
    Do While IsEmpty(hojaHorarios.Cells(filaPrimeraDatos, colCOD)) And _
             filaPrimeraDatos < hojaHorarios.Rows.Count
        filaPrimeraDatos = filaPrimeraDatos + 1
    Loop
    
    ' Las subcabeceras son todas las filas entre la cabecera principal y los datos
    filaSubcabeceraFin = filaPrimeraDatos - 1
    
    ' ============================
    '  COLUMNAS DE IDIOMA
    ' ============================
    ' Importante: si cambia el texto ("Inglés" -> "English") o se mueven,
    ' Find no las detecta y la macro no continúa.
    
    Set celdaTmp = hojaHorarios.Cells.Find("Inglés", LookIn:=xlValues, LookAt:=xlWhole)
    If celdaTmp Is Nothing Then Set celdaTmp = hojaHorarios.Cells.Find("Ingles", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celdaTmp Is Nothing Then colEN = celdaTmp.Column
    
    Set celdaTmp = hojaHorarios.Cells.Find("Español", LookIn:=xlValues, LookAt:=xlWhole)
    If celdaTmp Is Nothing Then Set celdaTmp = hojaHorarios.Cells.Find("Espanol", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celdaTmp Is Nothing Then colES = celdaTmp.Column
    
    Set celdaTmp = hojaHorarios.Cells.Find("Gallego", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celdaTmp Is Nothing Then colGL = celdaTmp.Column
    
    Set celdaTmp = hojaHorarios.Cells.Find("Catalán", LookIn:=xlValues, LookAt:=xlWhole)
    If celdaTmp Is Nothing Then Set celdaTmp = hojaHorarios.Cells.Find("Catalan", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celdaTmp Is Nothing Then colCA = celdaTmp.Column
    
    If colEN = 0 Or colES = 0 Then
        MsgBox "No se han encontrado correctamente las columnas de idioma.", vbCritical, "Horarios"
        Exit Sub
    End If
    
    ' ============================
    '  BLOQUES DE DÍAS
    ' ============================
    ' En esta plantilla normalmente existen:
    '   - Lunes a Viernes
    '   - Sábado
    '   - Domingo
    ' El código también soporta "Lunes a Jueves + Viernes" si aparece.
    colLunesJueves = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Lunes a Jueves")
    colViernes = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Viernes")
    colLunesViernes = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Lunes a Viernes")
    colSabado = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Sábado")
    colDomingo = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, "Domingo")
    
    ' ============================
    '  DÍAS ESPECIALES
    ' ============================
    ' Bloques completos con cabecera propia (igual que el horario habitual)
    nEsp = 0
    ReDim DiasEspeciales(1 To 1)
    
    AddDiaEspecial DiasEspeciales, nEsp, "MIE31", "Miércoles 31-12", "Miércoles 31"
    AddDiaEspecial DiasEspeciales, nEsp, "DOM04", "Domingo 04-01", "Domingo 04"
    AddDiaEspecial DiasEspeciales, nEsp, "LUN05", "Lunes 05-01", "Lunes 05"
    AddDiaEspecial DiasEspeciales, nEsp, "DOM11", "Domingo 11-01", "Domingo 11"
    
    ' Buscar dónde empieza cada bloque especial en la cabecera
    For i = 1 To nEsp
        With DiasEspeciales(i)
            .colInicio = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, .cabecera)
            If .colInicio = 0 And .cabeceraAlt <> "" Then
                .colInicio = BuscarColumnaPorTextoEnBloque(hojaHorarios, filaCabeceraPrincipal, filaSubcabeceraFin, .cabeceraAlt)
            End If
        End With
    Next i
    
    ' ============================
    '  FORMATO BASE DE LA HOJA
    ' ============================
    usaLunesJuevesMasViernes = (colLunesJueves <> 0 And colViernes <> 0)
    usaLunesViernes = (colLunesViernes <> 0)
    
    If Not usaLunesJuevesMasViernes And Not usaLunesViernes Then
        MsgBox "No se han encontrado correctamente los bloques de lunes.", vbCritical, "Horarios"
        Exit Sub
    End If
    
    If colSabado = 0 Then
        MsgBox "No se ha encontrado la columna de Sábado.", vbCritical, "Horarios"
        Exit Sub
    End If
    
    ' Calcular rangos de columnas (inicio/fin) para cada bloque.
    ' Se toma el siguiente bloque conocido como límite.
    If usaLunesJuevesMasViernes Then
        colLunesJuevesInicio = colLunesJueves
        colLunesJuevesFin = SiguienteInicioBloque(colLunesJuevesInicio, _
                                                  colViernes, colSabado, colDomingo, colEN, 0, 0, _
                                                  DiasEspeciales, nEsp) - 1
        
        colViernesInicio = colViernes
        colViernesFin = SiguienteInicioBloque(colViernesInicio, _
                                              colSabado, colDomingo, colEN, 0, 0, 0, _
                                              DiasEspeciales, nEsp) - 1
    ElseIf usaLunesViernes Then
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
    
    ' Para cada día especial, calcular su colFin igual que el resto (hasta el siguiente bloque)
    If nEsp > 0 Then
        For i = 1 To nEsp
            If DiasEspeciales(i).colInicio > 0 Then
                DiasEspeciales(i).colFin = SiguienteInicioBloque(DiasEspeciales(i).colInicio, _
                                                                 colEN, 0, 0, 0, 0, 0, _
                                                                 DiasEspeciales, nEsp) - 1
            End If
        Next i
    End If
    
    ' Detectar dentro de cada bloque las columnas de Apertura/Cierre
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
    
    ' Detectar Apertura/Cierre para días especiales (si existen en la hoja)
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
    
    ' Última fila con datos (por COD)
    ultimaFilaDatos = hojaHorarios.Cells(hojaHorarios.Rows.Count, colCOD).End(xlUp).Row
    
    ' Recorrer tiendas/filas y construir textos
    For filaActual = filaPrimeraDatos To ultimaFilaDatos
        
        ' Limpieza por iteración (para no arrastrar textos)
        extrasEN = ""
        extrasES = ""
        extrasGL = ""
        extrasCA = ""
        
        ' Lectura de horarios base según estructura de la hoja
        If usaLunesJuevesMasViernes Then
            horarioLunesJueves = LeerHorarioDeDia(hojaHorarios, filaActual, columnasLunesJueves)
            horarioViernes = LeerHorarioDeDia(hojaHorarios, filaActual, columnasViernes)
        ElseIf usaLunesViernes Then
            horarioLunesViernes = LeerHorarioDeDia(hojaHorarios, filaActual, columnasLunesViernes)
        End If
        
        horarioSabado = LeerHorarioDeDia(hojaHorarios, filaActual, columnasSabado)
        horarioDomingo = LeerHorarioDeDia(hojaHorarios, filaActual, columnasDomingo)
        
        ' Construcción del texto principal.
        ' Si viernes está vacío o es igual a L-J, se compacta a "Lun-Vie".
        If usaLunesJuevesMasViernes Then
            viernesVacio = (columnasViernes.NumTurnos = 0 Or EsDiaVacio(horarioViernes))
            viernesIgualLJ = (Not viernesVacio And SonHorariosIguales(horarioLunesJueves, horarioViernes))
            
            If viernesVacio Or viernesIgualLJ Then
                horarioLunesViernes = horarioLunesJueves
                casoHorario = ObtenerCasoHorario(horarioLunesViernes, horarioSabado, horarioDomingo)
                
                textoEN = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "EN")
                textoES = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "ES")
                textoGL = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "GL")
                textoCA = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "CA")
            Else
                textoEN = ConstruirTextoLunesJuevesViernes(horarioLunesJueves, horarioViernes, horarioSabado, horarioDomingo, "EN")
                textoES = ConstruirTextoLunesJuevesViernes(horarioLunesJueves, horarioViernes, horarioSabado, horarioDomingo, "ES")
                textoGL = ConstruirTextoLunesJuevesViernes(horarioLunesJueves, horarioViernes, horarioSabado, horarioDomingo, "GL")
                textoCA = ConstruirTextoLunesJuevesViernes(horarioLunesJueves, horarioViernes, horarioSabado, horarioDomingo, "CA")
            End If
        
        ElseIf usaLunesViernes Then
            casoHorario = ObtenerCasoHorario(horarioLunesViernes, horarioSabado, horarioDomingo)
            
            textoEN = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "EN")
            textoES = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "ES")
            textoGL = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "GL")
            textoCA = ConstruirTextoHorario(casoHorario, horarioLunesViernes, horarioSabado, horarioDomingo, "CA")
        End If
        
        ' Días especiales: se añaden en una línea aparte al final del texto
        If nEsp > 0 Then
            ReDim horEsp(1 To nEsp)
            
            For i = 1 To nEsp
                If DiasEspeciales(i).colInicio > 0 And DiasEspeciales(i).Columnas.NumTurnos > 0 Then
                    horEsp(i) = LeerHorarioDeDia(hojaHorarios, filaActual, DiasEspeciales(i).Columnas, True)
                Else
                    ' Si el bloque no existe en la hoja, dejamos el horario vacío
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
        
        ' Concatenar extras con salto de línea (Chr(10)) para que quede legible en celda
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
        
        ' Volcar resultados en la hoja. Formato texto para que Excel no "toquetee" el contenido.
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
    
    ' Corrección rápida de caracteres raros (típico de copias/CSV/encoding)
    With hojaHorarios.Cells
        .Replace What:="Ã¡", Replacement:="á", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã©", Replacement:="é", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã­", Replacement:="í", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã³", Replacement:="ó", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ãº", Replacement:="ú", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        
        .Replace What:="Ã", Replacement:="Á", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã‰", Replacement:="É", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã", Replacement:="Í", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã“", Replacement:="Ó", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ãš", Replacement:="Ú", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        
        .Replace What:="Ã±", Replacement:="ñ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ã‘", Replacement:="Ñ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        
        .Replace What:="Ã¼", Replacement:="ü", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="Ãœ", Replacement:="Ü", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    End With
    
    Exit Sub
ErrHandler:
    GestionarError NOMBRE_PROC
End Sub
