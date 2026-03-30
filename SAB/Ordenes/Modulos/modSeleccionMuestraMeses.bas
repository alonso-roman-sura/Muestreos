Option Explicit

' ========= ENTRADA DEL BOTÓN =========
Public Sub SeleccionMuestra()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("¿Está seguro de generar nuevas muestras por mes?", vbYesNo + vbQuestion, "Confirmar")
    If resp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo FIN
    Call SeleccionMuestra_PorMes

    MsgBox "Muestras por mes generadas correctamente.", vbInformation

FIN:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ========= LÓGICA PRINCIPAL =========
Private Sub SeleccionMuestra_PorMes()
    Dim w As Workbook: Set w = ThisWorkbook
    Dim wsM As Worksheet, wsO As Worksheet
    Set wsM = w.Worksheets("Muestra")
    Set wsO = w.Worksheets("Ordenes")

    Dim inicio As Range
    Set inicio = ResolveInicioMuestraCell(w)
    If inicio Is Nothing Then
        MsgBox "No encuentro una celda nombrada 'InicioMuestra' ni 'Inicio muestra'.", vbCritical
        Exit Sub
    End If

    ' Celda base de FORMATO: dos filas debajo de InicioMuestra (se usará para TODAS las celdas numéricas)
    Dim baseFmt As Range
    Set baseFmt = inicio.Offset(2, 0)

    ' === 1) Detectar meses presentes en Ordenes[Fecha] ===
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsO.ListObjects("Ordenes")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "No encuentro la tabla 'Ordenes'.", vbExclamation
        Exit Sub
    End If
    If lo.ListColumns("Fecha").DataBodyRange Is Nothing Then Exit Sub
    If WorksheetFunction.CountA(lo.ListColumns("Fecha").DataBodyRange) = 0 Then Exit Sub

    Dim arr, R As Long, dict As Object
    arr = lo.ListColumns("Fecha").DataBodyRange.Value
    Set dict = CreateObject("Scripting.Dictionary")

    For R = 1 To UBound(arr, 1)
        If IsDate(arr(R, 1)) Then
            Dim yy As Integer, mm As Integer
            yy = Year(arr(R, 1)): mm = Month(arr(R, 1))
            dict(yy & "-" & Format$(mm, "00")) = Array(yy, mm)
        End If
    Next R
    If dict.Count = 0 Then Exit Sub

    ' Ordenar claves YYYY-MM
    Dim keys() As String, i As Long, j As Long, tmp As String, k As Variant
    ReDim keys(0 To dict.Count - 1)
    i = 0
    For Each k In dict.keys
        keys(i) = CStr(k): i = i + 1
    Next k
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i

    ' === 1.b) Limpiar FORMATO antes de generar: desde baseFmt hasta el final usado de la hoja ===
    ClearFormatsFromAnchor wsM, baseFmt

    ' === 2) Por cada mes, escribir título y generar muestra ===
    Dim faltantes As Collection: Set faltantes = New Collection

    Dim idx As Long
    For idx = 0 To UBound(keys)
        Dim y0 As Long, m0 As Long
        y0 = dict(keys(idx))(0)
        m0 = dict(keys(idx))(1)

        Dim tag As String: tag = MesAbrevES_mSM(m0) & CStr(y0)     ' ej: Jul2025
        Dim nmUniv As String: nmUniv = "Universo" & tag
        Dim nmMues As String: nmMues = "Muestra" & tag

        Dim universo As Long, tamano As Long
        If Not NameToLongSafe(w, nmUniv, universo) Or Not NameToLongSafe(w, nmMues, tamano) Then
            faltantes.Add tag
            GoTo SiguienteMes
        End If
        If universo <= 0 Or tamano <= 0 Or tamano > universo Then
            faltantes.Add tag & " (universo/muestra inválidos)"
            GoTo SiguienteMes
        End If

        ' Bloque del mes: desplazar 6 columnas por mes respecto al inicio
        Dim colOffset As Long: colOffset = 6 * idx
        Dim celTitulo As Range: Set celTitulo = inicio.Offset(0, colOffset)
        Dim celInicioNums As Range: Set celInicioNums = inicio.Offset(2, colOffset) ' dos filas abajo

        ' 2.1) Título
        celTitulo.Value = "Muestra Mes " & (idx + 1) & " - " & MesAbrevES_mSM(m0) & " " & y0

        ' 2.2) Limpiar SOLO contenidos previos del bloque de números (5 columnas)
        ClearBlockContents wsM, celInicioNums, 5

        ' 2.3) Generar números únicos ordenados
        Dim nums() As Long
        nums = UniqueSortedSample(universo, tamano)

        ' 2.4) Escribir en grilla de 5 columnas, copiando formato SIEMPRE desde baseFmt
        WriteGrid wsM, celInicioNums, nums, 5, baseFmt

SiguienteMes:
    Next idx

    ' Reporte de faltantes (si hubiera)
    If faltantes.Count > 0 Then
        Dim msg As String: msg = "No se generaron estos meses por falta/valor inválido de nombres:" & vbCrLf
        For i = 1 To faltantes.Count
            msg = msg & "• " & CStr(faltantes(i)) & vbCrLf
        Next i
        MsgBox msg, vbExclamation
    End If
End Sub

' ========= HELPERS =========

' Resuelve la celda nombrada "InicioMuestra" o "Inicio muestra".
Private Function ResolveInicioMuestraCell(w As Workbook) As Range
    Dim nm As name
    On Error Resume Next
    Set nm = w.Names("InicioMuestra")
    If nm Is Nothing Then Set nm = w.Names("Inicio muestra")
    On Error GoTo 0
    If Not nm Is Nothing Then
        Set ResolveInicioMuestraCell = nm.RefersToRange
    Else
        Set ResolveInicioMuestraCell = Nothing
    End If
End Function

' Borra SOLO contenidos del bloque de datos (ancho = nCols) desde la celda de inicio hacia abajo.
Private Sub ClearBlockContents(ws As Worksheet, startCell As Range, ByVal nCols As Long)
    Dim c As Long, lastRow As Long, lastRowMax As Long
    lastRowMax = startCell.Row
    For c = 0 To nCols - 1
        lastRow = ws.Cells(ws.Rows.Count, startCell.Column + c).End(xlUp).Row
        If lastRow > lastRowMax Then lastRowMax = lastRow
    Next c
    If lastRowMax < startCell.Row Then Exit Sub
    ws.Range(startCell, ws.Cells(lastRowMax, startCell.Column + nCols - 1)).ClearContents
End Sub

' Limpia FORMATO desde la celda base hacia abajo y a la derecha,
' sin tocar la propia celda base (se usará como plantilla).
Private Sub ClearFormatsFromAnchor(ws As Worksheet, baseFmt As Range)
    Dim lastRow As Long, lastCol As Long

    With ws.UsedRange
        lastRow = .Row + .Rows.Count - 1
        lastCol = .Column + .Columns.Count - 1
    End With

    ' Derecha, misma fila (excluye la celda base)
    If baseFmt.Column + 1 <= lastCol Then
        ws.Range(ws.Cells(baseFmt.Row, baseFmt.Column + 1), _
                 ws.Cells(baseFmt.Row, lastCol)).ClearFormats
    End If

    ' Abajo, misma columna (excluye la celda base)
    If baseFmt.Row + 1 <= lastRow Then
        ws.Range(ws.Cells(baseFmt.Row + 1, baseFmt.Column), _
                 ws.Cells(lastRow, baseFmt.Column)).ClearFormats
    End If

    ' Cuadrante abajo-derecha (excluye la celda base)
    If baseFmt.Row + 1 <= lastRow And baseFmt.Column + 1 <= lastCol Then
        ws.Range(ws.Cells(baseFmt.Row + 1, baseFmt.Column + 1), _
                 ws.Cells(lastRow, lastCol)).ClearFormats
    End If
End Sub


' Escribe un vector en grilla de nCols columnas, copiando formatoBase a cada celda.
Private Sub WriteGrid(ws As Worksheet, startCell As Range, ByRef v() As Long, ByVal nCols As Long, formatBase As Range)
    Dim i As Long, R As Long, c As Long
    R = startCell.Row: c = startCell.Column
    For i = LBound(v) To UBound(v)
        ws.Cells(R, c).Value = v(i)
        formatBase.Copy
        ws.Cells(R, c).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        c = c + 1
        If c > startCell.Column + (nCols - 1) Then
            c = startCell.Column
            R = R + 1
        End If
    Next i
End Sub

' Muestra aleatoria de {1..universo} sin repetición, ordenada ascendentemente.
Private Function UniqueSortedSample(ByVal universo As Long, ByVal tamano As Long) As Long()
    Dim coll As Collection: Set coll = New Collection
    Dim rnum As Long
    Randomize
    Do While coll.Count < tamano
        rnum = Int(universo * Rnd) + 1
        On Error Resume Next
        coll.Add rnum, CStr(rnum)
        On Error GoTo 0
    Loop
    Dim v() As Long, i As Long, j As Long, tmp As Long
    ReDim v(1 To coll.Count)
    For i = 1 To coll.Count: v(i) = coll(i): Next i
    For i = 1 To UBound(v) - 1
        For j = i + 1 To UBound(v)
            If v(i) > v(j) Then tmp = v(i): v(i) = v(j): v(j) = tmp
        Next j
    Next i
    UniqueSortedSample = v
End Function

' Lee un nombre definido y lo convierte a Long de forma segura.
Private Function NameToLongSafe(w As Workbook, ByVal nm As String, ByRef outVal As Long) As Boolean
    Dim n As name
    On Error Resume Next
    Set n = w.Names(nm)
    On Error GoTo 0
    If n Is Nothing Then
        NameToLongSafe = False
        Exit Function
    End If
    Dim v As Variant: v = n.RefersToRange.Value
    If IsNumeric(v) Then
        outVal = CLng(v)
        NameToLongSafe = True
    Else
        NameToLongSafe = False
    End If
End Function

' Abreviaturas en español (privada y con nombre único para evitar ambigüedad).
Private Function MesAbrevES_mSM(ByVal m As Long) As String
    Select Case m
        Case 1: MesAbrevES_mSM = "Ene"
        Case 2: MesAbrevES_mSM = "Feb"
        Case 3: MesAbrevES_mSM = "Mar"
        Case 4: MesAbrevES_mSM = "Abr"
        Case 5: MesAbrevES_mSM = "May"
        Case 6: MesAbrevES_mSM = "Jun"
        Case 7: MesAbrevES_mSM = "Jul"
        Case 8: MesAbrevES_mSM = "Ago"
        Case 9: MesAbrevES_mSM = "Sep"
        Case 10: MesAbrevES_mSM = "Oct"
        Case 11: MesAbrevES_mSM = "Nov"
        Case 12: MesAbrevES_mSM = "Dic"
        Case Else: MesAbrevES_mSM = "Mes"
    End Select
End Function

