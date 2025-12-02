Option Explicit

'====================================================
'   TYPES
'====================================================

' Representa el horario de un día (hasta 2 turnos)
Type DayScheduleInfo
    StartTime1 As Variant
    EndTime1 As Variant
    StartTime2 As Variant
    EndTime2 As Variant
End Type

' Representa las columnas de apertura/cierre detectadas para un día
Type DayColumnInfo
    OpenCol(1 To 2) As Long
    CloseCol(1 To 2) As Long
    ShiftCount As Integer
End Type

'====================================================
'   GESTIÓN DE ERRORES CENTRALIZADA
'====================================================
Private Sub HandleError(ByVal procedureName As String)
    MsgBox "Error en procedimiento: " & procedureName & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, vbCritical
End Sub

'====================================================
'   FORMATEADOR DE HORAS
'====================================================
Private Function FormatTime(ByVal value As Variant) As String
    On Error GoTo ErrHandler
    
    If IsError(value) Or IsEmpty(value) Or value = "" Then
        FormatTime = ""
    Else
        FormatTime = Format$(value, "hh:mm")
    End If
    Exit Function

ErrHandler:
    FormatTime = ""
End Function

'====================================================
'   COMPARAR DOS DÍAS COMPLETOS
'====================================================
Private Function AreSchedulesEqual(d1 As DayScheduleInfo, d2 As DayScheduleInfo) As Boolean
    On Error GoTo ErrHandler
    
    AreSchedulesEqual = (FormatTime(d1.StartTime1) = FormatTime(d2.StartTime1) And _
                         FormatTime(d1.EndTime1) = FormatTime(d2.EndTime1) And _
                         FormatTime(d1.StartTime2) = FormatTime(d2.StartTime2) And _
                         FormatTime(d1.EndTime2) = FormatTime(d2.EndTime2))
    Exit Function
ErrHandler:
    AreSchedulesEqual = False
End Function

'====================================================
'   SABER SI UN DÍA ESTÁ VACÍO
'====================================================
Private Function IsDayEmpty(dayInfo As DayScheduleInfo) As Boolean
    On Error GoTo ErrHandler
    
    IsDayEmpty = (FormatTime(dayInfo.StartTime1) = "" And _
                  FormatTime(dayInfo.EndTime1) = "" And _
                  FormatTime(dayInfo.StartTime2) = "" And _
                  FormatTime(dayInfo.EndTime2) = "")
    Exit Function
ErrHandler:
    IsDayEmpty = True
End Function

'====================================================
'   TEXTO PARA UN DÍA
'====================================================
Private Function BuildDayText(dayInfo As DayScheduleInfo) As String
    On Error GoTo ErrHandler
    
    Dim txt As String
    
    If FormatTime(dayInfo.StartTime1) <> "" And FormatTime(dayInfo.EndTime1) <> "" Then
        txt = FormatTime(dayInfo.StartTime1) & " - " & FormatTime(dayInfo.EndTime1)
    End If
    
    If FormatTime(dayInfo.StartTime2) <> "" And FormatTime(dayInfo.EndTime2) <> "" Then
        If txt <> "" Then txt = txt & " / "
        txt = txt & FormatTime(dayInfo.StartTime2) & " - " & FormatTime(dayInfo.EndTime2)
    End If
    
    BuildDayText = txt
    Exit Function

ErrHandler:
    BuildDayText = ""
End Function

'====================================================
'   OBTENER CASO (1-2-0)
'====================================================
Private Function GetScheduleCase(monFri As DayScheduleInfo, sat As DayScheduleInfo, sun As DayScheduleInfo) As Integer
    On Error GoTo ErrHandler

    If AreSchedulesEqual(monFri, sat) And AreSchedulesEqual(monFri, sun) Then
        GetScheduleCase = 1
        Exit Function
    End If
    
    If AreSchedulesEqual(monFri, sat) And Not AreSchedulesEqual(monFri, sun) Then
        GetScheduleCase = 2
        Exit Function
    End If

    GetScheduleCase = 0
    Exit Function

ErrHandler:
    GetScheduleCase = 0
End Function

'====================================================
'   CONSTRUIR TEXTO COMPLETO SEGÚN CASOS
'====================================================
Private Function BuildScheduleText( _
    ByVal scheduleCase As Integer, _
    monFri As DayScheduleInfo, _
    sat As DayScheduleInfo, _
    sun As DayScheduleInfo, _
    ByVal lang As String) As String
    
    On Error GoTo ErrHandler
    
    Dim sMonFri As String, sMonSat As String, sMonSun As String
    Dim sSat As String, sSun As String
    Dim sep As String: sep = " | "
    Dim result As String, part As String

    ' Prefijos según idioma
    Select Case UCase(lang)
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
    
    Select Case scheduleCase
    
        Case 1
            BuildScheduleText = sMonSun & BuildDayText(monFri)
        
        Case 2
            If IsDayEmpty(sun) Then
                BuildScheduleText = sMonSat & BuildDayText(monFri)
            Else
                BuildScheduleText = sMonSat & BuildDayText(monFri) & _
                                    sep & sSun & BuildDayText(sun)
            End If
        
        Case Else
            result = ""
            
            If Not IsDayEmpty(monFri) Then
                part = sMonFri & BuildDayText(monFri)
                If result <> "" Then result = result & sep
                result = result & part
            End If
            
            If Not IsDayEmpty(sat) Then
                part = sSat & BuildDayText(sat)
                If result <> "" Then result = result & sep
                result = result & part
            End If
            
            If Not IsDayEmpty(sun) Then
                part = sSun & BuildDayText(sun)
                If result <> "" Then result = result & sep
                result = result & part
            End If
            
            BuildScheduleText = result
    End Select

    Exit Function

ErrHandler:
    BuildScheduleText = ""
End Function

'====================================================
'   DOMINGO 30 TEXTO
'====================================================
Public Function SpecialSundayText(lang As String) As String
    On Error GoTo ErrHandler
    
    Select Case UCase(lang)
        Case "EN": SpecialSundayText = "Sunday Nov 30: "
        Case "ES": SpecialSundayText = "Domingo 30 Nov: "
        Case "GL": SpecialSundayText = "Domingo 30 Nov: "
        Case "CA": SpecialSundayText = "Diumenge 30 Nov: "
        Case Else: SpecialSundayText = "Sunday Nov 30: "
    End Select
    
    Exit Function
ErrHandler:
    SpecialSundayText = ""
End Function

'====================================================
'   HELPERS PARA BUSCAR COLUMNAS
'====================================================
Private Function FindColumnInRowByTitle( _
    ws As Worksheet, _
    ByVal row As Long, _
    ByVal title As String) As Long
    
    On Error GoTo ErrHandler
    
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(row, c).Value)) = title Then
            FindColumnInRowByTitle = c
            Exit Function
        End If
    Next c
    Exit Function

ErrHandler:
    FindColumnInRowByTitle = 0
End Function

Private Function FindColumnByTextInBlock( _
    ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long, _
    ByVal title As String) As Long
    
    On Error GoTo ErrHandler
    
    Dim lastCol As Long, r As Long, c As Long
    lastCol = ws.Cells(rowStart, ws.Columns.Count).End(xlToLeft).Column
    
    For r = rowStart To rowEnd
        For c = 1 To lastCol
            If Trim$(CStr(ws.Cells(r, c).Value)) = title Then
                FindColumnByTextInBlock = c
                Exit Function
            End If
        Next c
    Next r
    Exit Function

ErrHandler:
    FindColumnByTextInBlock = 0
End Function

'====================================================
'   DETECTAR COLS APERTURA/CIERRE
'====================================================
Private Function DetectDayColumns( _
    ws As Worksheet, _
    ByVal subHeaderStart As Long, _
    ByVal subHeaderEnd As Long, _
    ByVal colStart As Long, _
    ByVal colEnd As Long) As DayColumnInfo
    
    On Error GoTo ErrHandler
    
    Dim dc As DayColumnInfo
    Dim c As Long, r As Long
    Dim txt As String
    Dim idx As Integer
    Dim completePairs As Integer
    
    If colStart = 0 Or colEnd = 0 Then
        DetectDayColumns = dc
        Exit Function
    End If
    
    For c = colStart To colEnd
        For r = subHeaderStart To subHeaderEnd
        
            txt = Trim$(CStr(ws.Cells(r, c).Value))
            
            If txt <> "" Then
            
                If StrComp(txt, "Apertura", vbTextCompare) = 0 Then
                    If dc.ShiftCount < 2 Then
                        dc.ShiftCount = dc.ShiftCount + 1
                        dc.OpenCol(dc.ShiftCount) = c
                    End If
                    Exit For
                
                ElseIf StrComp(txt, "Cierre", vbTextCompare) = 0 Then
                    For idx = 1 To 2
                        If dc.CloseCol(idx) = 0 Then
                            dc.CloseCol(idx) = c
                            Exit For
                        End If
                    Next idx
                    Exit For
                End If
            
            End If
        
        Next r
    Next c
    
    completePairs = 0
    For idx = 1 To 2
        If dc.OpenCol(idx) <> 0 And dc.CloseCol(idx) <> 0 Then completePairs = completePairs + 1
    Next idx
    
    dc.ShiftCount = completePairs
    DetectDayColumns = dc
    Exit Function

ErrHandler:
    DetectDayColumns = dc
End Function

'====================================================
'   LEER HORARIO DE UN DÍA
'====================================================
Private Function ReadDaySchedule( _
    ws As Worksheet, _
    ByVal row As Long, _
    dayCols As DayColumnInfo) As DayScheduleInfo
    
    On Error GoTo ErrHandler
    
    Dim ds As DayScheduleInfo
    Dim A1 As Variant, C1 As Variant
    Dim A2 As Variant, C2 As Variant
    Dim start1 As Variant, end1 As Variant
    Dim start2 As Variant, end2 As Variant
    
    If dayCols.ShiftCount = 0 Then
        ReadDaySchedule = ds
        Exit Function
    End If
    
    If dayCols.OpenCol(1) <> 0 Then A1 = ws.Cells(row, dayCols.OpenCol(1)).Value
    If dayCols.CloseCol(1) <> 0 Then C1 = ws.Cells(row, dayCols.CloseCol(1)).Value
    
    If dayCols.ShiftCount >= 2 Then
        If dayCols.OpenCol(2) <> 0 Then A2 = ws.Cells(row, dayCols.OpenCol(2)).Value
        If dayCols.CloseCol(2) <> 0 Then C2 = ws.Cells(row, dayCols.CloseCol(2)).Value
    End If
    
    If FormatTime(A1) <> "" And FormatTime(C1) <> "" And _
       FormatTime(A2) <> "" And FormatTime(C2) <> "" Then
       
        start1 = A1: end1 = C1
        start2 = A2: end2 = C2
    
    Else
        
        If FormatTime(A1) <> "" Then start1 = A1
        If FormatTime(A2) <> "" Then
            If IsEmpty(start1) Or (FormatTime(A2) <> "" And A2 < start1) Then start1 = A2
        End If
        
        If FormatTime(C1) <> "" Then end1 = C1
        If FormatTime(C2) <> "" Then
            If IsEmpty(end1) Or (FormatTime(C2) <> "" And C2 > end1) Then end1 = C2
        End If
    End If
    
    ds.StartTime1 = start1
    ds.EndTime1 = end1
    ds.StartTime2 = start2
    ds.EndTime2 = end2
    
    ReadDaySchedule = ds
    Exit Function

ErrHandler:
    ReadDaySchedule = ds
End Function

'====================================================
'   MACRO PRINCIPAL
'====================================================
Public Sub Horarios()

    Const PROC As String = "Horarios"
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim subHeaderStart As Long, subHeaderEnd As Long
    Dim firstDataRow As Long, lastDataRow As Long
    
    Dim colCode As Long
    Dim colMonFri As Long, colSat As Long, colSun As Long, colSun30 As Long
    Dim colEN As Long, colES As Long, colGL As Long, colCA As Long
    
    Dim monFriCols As DayColumnInfo, satCols As DayColumnInfo
    Dim sunCols As DayColumnInfo, sun30Cols As DayColumnInfo
    
    Dim monFri As DayScheduleInfo
    Dim sat As DayScheduleInfo
    Dim sun As DayScheduleInfo
    Dim sun30 As DayScheduleInfo
    
    Dim scheduleCase As Integer
    
    Dim textEN As String, textES As String, textGL As String, textCA As String
    
    Dim cell As Range
    
    '------------------------------------------------
    ' Buscar hoja
    '------------------------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Horarios habituales")
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("HORARIO ESPAÑA")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de horarios.", vbCritical
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Localizar "COD"
    '------------------------------------------------
    Set cell = ws.Cells.Find("COD", LookIn:=xlValues, LookAt:=xlWhole)
    
    If cell Is Nothing Then
        MsgBox "No se encontró la cabecera 'COD'.", vbCritical
        Exit Sub
    End If
    
    headerRow = cell.Row
    colCode = cell.Column
    
    subHeaderStart = headerRow + 1
    
    firstDataRow = headerRow + 1
    Do While IsEmpty(ws.Cells(firstDataRow, colCode)) And firstDataRow < ws.Rows.Count
        firstDataRow = firstDataRow + 1
    Loop
    
    subHeaderEnd = firstDataRow - 1
    
    '------------------------------------------------
    ' Columnas de idiomas
    '------------------------------------------------
    colEN = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Inglés")
    colES = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Español")
    colGL = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Gallego")
    colCA = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Catalán")
    
    If colEN = 0 Then colEN = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Ingles")
    If colES = 0 Then colES = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Espanol")
    If colCA = 0 Then colCA = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Catalan")
    
    If colEN = 0 Or colES = 0 Then
        MsgBox "No se encontraron columnas de idioma.", vbCritical
        Exit Sub
    End If
    
    '------------------------------------------------
    ' Columnas de días
    '------------------------------------------------
    colMonFri = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Lunes a Viernes")
    colSat = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Sábado")
    colSun = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Domingo")
    colSun30 = FindColumnByTextInBlock(ws, headerRow, subHeaderEnd, "Domingo 30")
    
    '------------------------------------------------
    ' Rango de días
    '------------------------------------------------
    Dim colMonFriStart As Long, colMonFriEnd As Long
    Dim colSatStart As Long, colSatEnd As Long
    Dim colSunStart As Long, colSunEnd As Long
    Dim colSun30Start As Long, colSun30End As Long
    
    colMonFriStart = colMonFri
    colSatStart = colSat
    colSunStart = colSun
    colSun30Start = colSun30
    
    colMonFriEnd = colSatStart - 1
    colSatEnd = colSunStart - 1
    
    If colSun30 > 0 Then
        colSunEnd = colSun30Start - 1
        colSun30End = colEN - 1
    Else
        colSunEnd = colEN - 1
    End If
    
    '------------------------------------------------
    ' Detectar columnas de apertura/cierre
    '------------------------------------------------
    monFriCols = DetectDayColumns(ws, subHeaderStart, subHeaderEnd, colMonFriStart, colMonFriEnd)
    satCols = DetectDayColumns(ws, subHeaderStart, subHeaderEnd, colSatStart, colSatEnd)
    sunCols = DetectDayColumns(ws, subHeaderStart, subHeaderEnd, colSunStart, colSunEnd)
    If colSun30 > 0 Then
        sun30Cols = DetectDayColumns(ws, subHeaderStart, subHeaderEnd, colSun30Start, colSun30End)
    End If
    
    '------------------------------------------------
    ' Última fila
    '------------------------------------------------
    lastDataRow = ws.Cells(ws.Rows.Count, colCode).End(xlUp).Row
    
    '------------------------------------------------
    ' Bucle principal
    '------------------------------------------------
    Dim row As Long
    
    For row = firstDataRow To lastDataRow
        
        monFri = ReadDaySchedule(ws, row, monFriCols)
        sat = ReadDaySchedule(ws, row, satCols)
        sun = ReadDaySchedule(ws, row, sunCols)
        
        If colSun30 > 0 Then
            sun30 = ReadDaySchedule(ws, row, sun30Cols)
        Else
            With sun30
                .StartTime1 = Empty: .EndTime1 = Empty
                .StartTime2 = Empty: .EndTime2 = Empty
            End With
        End If
        
        scheduleCase = GetScheduleCase(monFri, sat, sun)
        
        textEN = BuildScheduleText(scheduleCase, monFri, sat, sun, "EN")
        textES = BuildScheduleText(scheduleCase, monFri, sat, sun, "ES")
        textGL = BuildScheduleText(scheduleCase, monFri, sat, sun, "GL")
        textCA = BuildScheduleText(scheduleCase, monFri, sat, sun, "CA")
        
        If colSun30 > 0 And Not IsDayEmpty(sun30) Then
            
            If textEN <> "" Then textEN = textEN & Chr(10)
            If textES <> "" Then textES = textES & Chr(10)
            If textGL <> "" Then textGL = textGL & Chr(10)
            If textCA <> "" Then textCA = textCA & Chr(10)
            
            textEN = textEN & SpecialSundayText("EN") & BuildDayText(sun30)
            textES = textES & SpecialSundayText("ES") & BuildDayText(sun30)
            textGL = textGL & SpecialSundayText("GL") & BuildDayText(sun30)
            textCA = textCA & SpecialSundayText("CA") & BuildDayText(sun30)
        End If
        
        ws.Cells(row, colEN).NumberFormat = "@": ws.Cells(row, colEN).Value = textEN
        ws.Cells(row, colES).NumberFormat = "@": ws.Cells(row, colES).Value = textES
        If colGL > 0 Then ws.Cells(row, colGL).NumberFormat = "@": ws.Cells(row, colGL).Value = textGL
        If colCA > 0 Then ws.Cells(row, colCA).NumberFormat = "@": ws.Cells(row, colCA).Value = textCA
    
    Next row
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
