' ========== modExportarMuestra.bas ==========
Option Explicit

' BOTÓN: Generar tablas por mes desde los números de muestra
Public Sub GenerarTablas()
    Dim wsO As Worksheet, lo As ListObject
    On Error Resume Next
    Set wsO = ThisWorkbook.Worksheets("Ordenes")
    If wsO Is Nothing Then
        MsgBox "No existe la hoja 'Ordenes'.", vbExclamation: Exit Sub
    End If
    Set lo = wsO.ListObjects("Ordenes")
    If lo Is Nothing Then If wsO.ListObjects.Count > 0 Then Set lo = wsO.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "No encontré la tabla 'Ordenes'.", vbExclamation: Exit Sub
    End If

    Dim ancla As Range
    On Error Resume Next
    Set ancla = ThisWorkbook.Names("InicioMuestra").RefersToRange
    On Error GoTo 0
    If ancla Is Nothing Then
        MsgBox "Falta el nombre definido 'InicioMuestra'.", vbExclamation: Exit Sub
    End If

    If IsEmpty(ancla.Value) Or Len(Trim$(CStr(ancla.Value))) = 0 Then
        MsgBox "No se han generado los números de muestra." & vbCrLf & _
               "Primero ejecute 'Seleccionar Muestras'.", vbExclamation, "Sin muestra"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim k As Long, hechos As Long
    k = 0
    Do
        Dim hdr As Range, lbl As String, y As Long, m As Long
        Set hdr = ancla.offset(0, 6 * k)
        lbl = CStr(hdr.Value)

        If Len(lbl) = 0 Then Exit Do
        If Left$(lbl, 12) <> "Muestra Mes " Then Exit Do
        If Not ParseEtiquetaMes(lbl, y, m) Then Exit Do

        ' Lee números de muestra debajo (dos filas abajo) en bloque de 5 columnas
        Dim nums() As Long
        nums = LeerNumerosDeMuestra(ancla.offset(2, 6 * k))
        If UBoundSafe(nums) = 0 Then GoTo SiguienteMes

        ' Filtra filas del mes y ordénalas por Fecha y Hora (y descarta NºOrden vacío)
        Dim idxMes() As Long
        idxMes = FilasDelMes_Ordenadas(lo, y, m)
        If UBoundSafe(idxMes) = 0 Then GoTo SiguienteMes

        ' Selecciona por posición dentro del mes
        Dim selIdx() As Long, i As Long, c As Long
        ReDim selIdx(1 To UBoundSafe(nums))
        c = 0
        For i = 1 To UBoundSafe(nums)
            If nums(i) >= 1 And nums(i) <= UBoundSafe(idxMes) Then
                c = c + 1
                selIdx(c) = idxMes(nums(i))
            End If
        Next i
        If c = 0 Then GoTo SiguienteMes
        ReDim Preserve selIdx(1 To c)

        ' Exportar
        Dim nombreHoja As String, nombreTabla As String
        nombreHoja = "Ordenes_Muestra_" & MesAbrevES(m) & "_" & CStr(y)
        nombreTabla = nombreHoja

        ExportarFilas lo, selIdx, nombreHoja, nombreTabla
        hechos = hechos + 1

SiguienteMes:
        k = k + 1
    Loop

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Exportación finalizada. Hojas creadas: " & hechos & ".", vbInformation
End Sub

' --------- Lee números en columnas (5 de ancho) desde celda inicio hacia abajo ----------
Private Function LeerNumerosDeMuestra(startCell As Range) As Long()
    Dim nums() As Long, cap As Long, R As Long, c As Long, v
    cap = 0
    ReDim nums(0 To 0)

    R = 0
    Do
        Dim filaVacia As Boolean: filaVacia = True
        For c = 0 To 4
            v = startCell.offset(R, c).Value
            If Len(v) > 0 Then
                filaVacia = False
                If IsNumeric(v) Then
                    cap = cap + 1
                    ReDim Preserve nums(0 To cap)
                    nums(cap) = CLng(v)
                End If
            End If
        Next c
        If filaVacia Then Exit Do
        R = R + 1
    Loop

    LeerNumerosDeMuestra = nums
End Function

Private Function UBoundSafe(arr As Variant) As Long
    On Error GoTo fallo
    UBoundSafe = UBound(arr)
    Exit Function
fallo:
    UBoundSafe = 0
End Function

' --- Convierte cualquier variante a fracción de día segura (00:00–<1)
Private Function TimeFrac(ByVal v As Variant) As Double
    On Error GoTo fallback
    If IsDate(v) Then
        TimeFrac = CDbl(CDate(v)) - Fix(CDbl(CDate(v)))
        Exit Function
    ElseIf IsNumeric(v) Then
        TimeFrac = CDbl(v) - Fix(CDbl(v))
        If TimeFrac < 0 Then TimeFrac = 0
        Exit Function
    End If
fallback:
    On Error Resume Next
    TimeFrac = CDbl(TimeValue(CStr(v)))
    If Err.Number <> 0 Then TimeFrac = 0
    On Error GoTo 0
End Function

' --------- Devuelve índices (filas absolutas del DataBodyRange) del mes, ordenados Fecha+Hora ----------
' Excluye registros con NºOrden vacío, para que el universo mensual coincida con los números generados.
Private Function FilasDelMes_Ordenadas(lo As ListObject, ByVal y As Long, ByVal m As Long) As Long()
    Dim i As Long, n As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    n = lo.DataBodyRange.Rows.Count

    Dim cF As Long, cH As Long, cN As Long
    On Error Resume Next
    cF = lo.ListColumns("Fecha").Index
    cH = lo.ListColumns("Hora").Index
    cN = lo.ListColumns("NºOrden").Index
    On Error GoTo 0
    If cF = 0 Or cH = 0 Or cN = 0 Then Exit Function

    Dim idx() As Long, key() As Double, k As Long
    ReDim idx(1 To n)
    ReDim key(1 To n)

    Dim R As Range, d As Variant, t As Variant, nro As Variant
    For i = 1 To n
        Set R = lo.DataBodyRange.Rows(i)
        d = R.Cells(1, cF).Value
        t = R.Cells(1, cH).Value
        nro = R.Cells(1, cN).Value

        If IsDate(d) Then
            If Year(d) = y And Month(d) = m Then
                If Len(Trim$(CStr(nro))) > 0 Then
                    k = k + 1
                    idx(k) = i
                    key(k) = CDbl(CDate(d)) + TimeFrac(t)
                End If
            End If
        End If
    Next i

    If k = 0 Then Exit Function
    ReDim Preserve idx(1 To k)
    ReDim Preserve key(1 To k)

    QuickSortKeys key, idx, 1, k
    FilasDelMes_Ordenadas = idx
End Function

' --------- Exporta filas seleccionadas a hoja nueva como tabla; formatea Fecha/Hora ----------
Private Sub ExportarFilas(lo As ListObject, ByRef selIdx() As Long, _
                          ByVal sheetName As String, ByVal tableName As String)

    ' Crear/limpiar hoja destino
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then ws.Delete
    Set ws = Nothing
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName

    ' Encabezados
    Dim headers As Variant
    headers = lo.headerRowRange.Value
    ws.Range("A1").Resize(1, UBound(headers, 2)).Value = headers

    ' Volcar filas
    Dim i As Long, dstRow As Long
    dstRow = 2
    For i = 1 To UBound(selIdx)
        ws.Range("A" & dstRow).Resize(1, lo.ListColumns.Count).Value = _
            lo.DataBodyRange.Rows(selIdx(i)).Value
        dstRow = dstRow + 1
    Next i

    ' Crear tabla
    Dim loT As ListObject
    Set loT = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    loT.Name = tableName

    ' Estilo de tabla: Claro 9 (TableStyleLight9). Intento en inglés y, si falla, en español.
    On Error Resume Next
    loT.TableStyle = "TableStyleLight9"
    If loT.TableStyle <> "TableStyleLight9" Then loT.TableStyle = "Claro 9"
    On Error GoTo 0

    ' Formatos: Fecha como fecha y Hora como hora larga
    On Error Resume Next
    Dim cF As Long, cH As Long
    cF = loT.ListColumns("Fecha").Index
    cH = loT.ListColumns("Hora").Index
    On Error GoTo 0

    If cF > 0 Then loT.ListColumns(cF).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"
    If cH > 0 Then loT.ListColumns(cH).DataBodyRange.NumberFormatLocal = "hh:mm:ss"

    ' AutoFit solo para el rango de la tabla
    loT.Range.Columns.AutoFit
End Sub


' --------- Parsear "Muestra Mes X - Jul 2025" -> (y,m) ----------
Private Function ParseEtiquetaMes(ByVal s As String, ByRef y As Long, ByRef m As Long) As Boolean
    Dim p As Long: p = InStr(1, s, " - ", vbTextCompare)
    If p = 0 Then Exit Function
    Dim tail As String: tail = Trim$(Mid$(s, p + 3))
    Dim parts() As String: parts = Split(tail, " ")
    If UBound(parts) <> 1 Then Exit Function
    y = val(parts(1))
    m = MesNumeroAbrev(parts(0))
    ParseEtiquetaMes = (m >= 1 And y > 0)
End Function

Private Function MesAbrevES(ByVal m As Long) As String
    Select Case m
        Case 1: MesAbrevES = "Ene"
        Case 2: MesAbrevES = "Feb"
        Case 3: MesAbrevES = "Mar"
        Case 4: MesAbrevES = "Abr"
        Case 5: MesAbrevES = "May"
        Case 6: MesAbrevES = "Jun"
        Case 7: MesAbrevES = "Jul"
        Case 8: MesAbrevES = "Ago"
        Case 9: MesAbrevES = "Sep"
        Case 10: MesAbrevES = "Oct"
        Case 11: MesAbrevES = "Nov"
        Case 12: MesAbrevES = "Dic"
        Case Else: MesAbrevES = "Mes"
    End Select
End Function

Private Function MesNumeroAbrev(ByVal s As String) As Long
    Dim u As String: u = LCase$(Trim$(s))
    Select Case u
        Case "ene": MesNumeroAbrev = 1
        Case "feb": MesNumeroAbrev = 2
        Case "mar": MesNumeroAbrev = 3
        Case "abr": MesNumeroAbrev = 4
        Case "may": MesNumeroAbrev = 5
        Case "jun": MesNumeroAbrev = 6
        Case "jul": MesNumeroAbrev = 7
        Case "ago": MesNumeroAbrev = 8
        Case "sep", "set": MesNumeroAbrev = 9
        Case "oct": MesNumeroAbrev = 10
        Case "nov": MesNumeroAbrev = 11
        Case "dic": MesNumeroAbrev = 12
        Case Else: MesNumeroAbrev = 0
    End Select
End Function

' --------- Ordenación rápida por clave numérica ascendente ----------
Private Sub QuickSortKeys(ByRef key() As Double, ByRef idx() As Long, ByVal L As Long, ByVal R As Long)
    Dim i As Long, j As Long, p As Double, tD As Double, tL As Long
    i = L: j = R: p = key((L + R) \ 2)
    Do While i <= j
        Do While key(i) < p: i = i + 1: Loop
        Do While key(j) > p: j = j - 1: Loop
        If i <= j Then
            tD = key(i): key(i) = key(j): key(j) = tD
            tL = idx(i): idx(i) = idx(j): idx(j) = tL
            i = i + 1: j = j - 1
        End If
    Loop
    If L < j Then QuickSortKeys key, idx, L, j
    If i < R Then QuickSortKeys key, idx, i, R
End Sub