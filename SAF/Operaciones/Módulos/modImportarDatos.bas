' ========== modImportarDatos.bas ==========
Option Explicit

Private Const DEBUG_IMPORT As Boolean = False

' Columnas esperadas (23) - "Pocentaje" con typo es el nombre real del sistema
' Se acepta también "Porcentaje" como failsafe
Private Const N_COLS       As Long = 23
Private Const COL_OPERACION As Long = 13   ' posición de la columna Operación (1-based)
Private Const COL_FECHA_OP  As Long = 3    ' posición de Fecha de Operación (1-based)

Private Sub InicializarHeaders(ByRef hdrs() As String)
    ReDim hdrs(1 To N_COLS)
    hdrs(1) = "Portafolio"
    hdrs(2) = "Codigo de Orden"
    hdrs(3) = "Fecha de Operaci" & Chr(243) & "n"
    hdrs(4) = "Fecha Liquidacion"
    hdrs(5) = "Fecha fin Contrato"
    hdrs(6) = "Codigo ISIN"
    hdrs(7) = "Codigo SBS"
    hdrs(8) = "Monto de Operaci" & Chr(243) & "n Original"
    hdrs(9) = "Monto de Operaci" & Chr(243) & "n ML"
    hdrs(10) = "Cantidad"
    hdrs(11) = "Precio"
    hdrs(12) = "Codigo de Emisor"
    hdrs(13) = "Operaci" & Chr(243) & "n"
    hdrs(14) = "Moneda"
    hdrs(15) = "Nemonico"
    hdrs(16) = "Codigo de Tercero"
    hdrs(17) = "Tercero"
    hdrs(18) = "Monto Nominal Operaci" & Chr(243) & "n Original"
    hdrs(19) = "Monto Nominal Operaci" & Chr(243) & "n ML"
    hdrs(20) = "Total de Comisiones"
    hdrs(21) = "Plaza"
    hdrs(22) = "Tipo Tasa"
    hdrs(23) = "Pocentaje Tasa"   ' typo del sistema; failsafe en CanonHeader
End Sub

' ============================================================
'  ENTRADA DEL BOTÓN "Importar Datos"
' ============================================================
Public Sub ImportarDatos()
    Dim ruta As String
    ruta = PickFilePath()
    If Len(ruta) = 0 Then Exit Sub

    Dim paso As String
    Dim wbOrigen As Workbook
    On Error GoTo ERR_HANDLER

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    paso = "Abrir workbook"
    Set wbOrigen = Workbooks.Open(Filename:=ruta, ReadOnly:=True, _
                                  UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)

    If EsTextoOHtml(wbOrigen) Then
        wbOrigen.Close False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "El archivo seleccionado no es un libro Excel nativo.", vbExclamation
        Exit Sub
    End If

    paso = "Leer hoja origen"
    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Sheets(1)

    paso = "Validar encabezados"
    Dim headerRow As Long
    Dim hdrs() As String
    InicializarHeaders hdrs

    If Not ValidarEncabezados(wsOrigen, hdrs, headerRow) Then
        wbOrigen.Close False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "El archivo no tiene el formato de Operaciones SAF esperado." & vbCrLf & _
               "Se buscaron los 23 encabezados en filas 1 a 5.", _
               vbCritical, "Formato no reconocido"
        Exit Sub
    End If

    paso = "Calcular rango de datos"
    Dim lastCol As Long
    lastCol = wsOrigen.Cells(headerRow, wsOrigen.Columns.Count).End(xlToLeft).Column
    If lastCol < N_COLS Then lastCol = N_COLS

    Dim lastRow As Long
    lastRow = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    If lastRow <= headerRow Then
        wbOrigen.Close False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "No se encontraron filas de datos debajo del encabezado.", vbExclamation
        Exit Sub
    End If

    Dim rngDatos As Range
    Set rngDatos = wsOrigen.Range( _
        wsOrigen.Cells(headerRow, 1), _
        wsOrigen.Cells(lastRow, lastCol))

    paso = "Preparar hoja Operaciones"
    Dim wsDest As Worksheet
    PrepararHojaOperaciones wsDest

    ' Formatear columnas de fecha como texto ANTES de pegar
    ' para que el serial de Excel no sea malinterpretado por el locale del sistema
    paso = "Formatear columnas de fecha como texto"
    Dim fechaCols As Variant
    fechaCols = Array(COL_FECHA_OP, 4, 5)
    Dim fc As Variant
    For Each fc In fechaCols
        wsDest.Columns(CLng(fc)).NumberFormat = "@"
    Next fc

    paso = "Copiar valores (" & rngDatos.Rows.Count & " filas x " & rngDatos.Columns.Count & " cols)"
    wsDest.Range("A1").Resize(rngDatos.Rows.Count, rngDatos.Columns.Count).Value = rngDatos.Value

    ' Re-leer cada celda de fecha como .Text desde el origen
    ' para preservar exactamente DD/MM/YYYY tal como lo muestra Excel en el archivo fuente
    paso = "Preservar formato DD/MM/YYYY en fechas"
    For Each fc In fechaCols
        Dim srcFechaCol As Range
        Set srcFechaCol = rngDatos.Columns(CLng(fc))
        Dim rr As Long
        For rr = 2 To rngDatos.Rows.Count
            Dim srcFechaTxt As String
            srcFechaTxt = srcFechaCol.Cells(rr, 1).Text
            If Len(Trim$(srcFechaTxt)) > 0 Then
                wsDest.Cells(rr, CLng(fc)).Value = srcFechaTxt
            End If
        Next rr
    Next fc

    paso = "Normalizar encabezado col 23 (Pocentaje/Porcentaje)"
    ' Normalizar el header de col 23 al valor canónico "Pocentaje Tasa"
    ' para que PQ lo encuentre de forma consistente
    Dim hdr23 As String: hdr23 = Trim$(CStr(wsDest.Cells(1, 23).Value))
    If InStr(UCase$(hdr23), "PORCENTAJE") > 0 Or InStr(UCase$(hdr23), "POCENTAJE") > 0 Then
        wsDest.Cells(1, 23).Value = "Pocentaje Tasa"
    End If

    paso = "Cerrar workbook origen"
    wbOrigen.Close False
    Set wbOrigen = Nothing

    paso = "Crear consulta Power Query"
    UpsertQuery "Operaciones", M_Operaciones_PQ()

    paso = "Crear tabla Operaciones_Raw"
    Dim loRaw As ListObject
    On Error Resume Next
    Set loRaw = wsDest.ListObjects("Operaciones_Raw")
    If Not loRaw Is Nothing Then loRaw.Unlist: Set loRaw = Nothing
    On Error GoTo ERR_HANDLER
    Set loRaw = wsDest.ListObjects.Add(xlSrcRange, wsDest.Range("A1").CurrentRegion, , xlYes)
    loRaw.name = "Operaciones_Raw"

    paso = "Cargar consulta PQ en hoja Operaciones"
    CargarConsultaEnHoja "Operaciones", "Operaciones", "Operaciones"

    paso = "Formato fecha y autofit"
    Dim wsOpFinal As Worksheet
    On Error Resume Next
    Set wsOpFinal = ThisWorkbook.Worksheets("Operaciones")
    On Error GoTo ERR_HANDLER
    If Not wsOpFinal Is Nothing Then
        Dim loTablaFinal As ListObject
        On Error Resume Next
        Set loTablaFinal = wsOpFinal.ListObjects("Operaciones")
        On Error GoTo ERR_HANDLER
        If Not loTablaFinal Is Nothing Then
            AplicarFormatosFecha loTablaFinal
            loTablaFinal.Range.Columns.AutoFit
        End If
    End If

    If DEBUG_IMPORT Then GenerarHojaDebug wsDest, ruta

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    paso = "Autodetectar per" & Chr(237) & "odo"
    AutodetectarPeriodo

    paso = "Tama" & Chr(241) & "oPoblacion"
    TamañoPoblacion

    MsgBox "Operaciones SAF importadas correctamente.", vbInformation
    Exit Sub

ERR_HANDLER:
    Dim errDesc As String: errDesc = "[" & paso & "] " & Err.Description
    On Error Resume Next
    If Not wbOrigen Is Nothing Then wbOrigen.Close False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error al importar:" & vbCrLf & vbCrLf & errDesc, _
           vbCritical, "Error de importaci" & Chr(243) & "n"
End Sub

' ============================================================
'  VALIDAR ENCABEZADOS
'  Busca la fila (1..5) donde aparezcan al menos 20 de los 23
'  encabezados esperados, usando comparación canonizada.
'  Acepta "Pocentaje" y "Porcentaje" como equivalentes.
' ============================================================
Private Function ValidarEncabezados(ws As Worksheet, _
                                     ByRef hdrs() As String, _
                                     ByRef headerRow As Long) As Boolean
    Dim R As Long, c As Long, matches As Long
    For R = 1 To 5
        matches = 0
        For c = 1 To Application.Min(N_COLS + 3, 30)
            Dim cv As String: cv = CanonHeader(CStr(ws.Cells(R, c).Value))
            Dim h As Long
            For h = 1 To N_COLS
                If cv = CanonHeader(hdrs(h)) Then
                    matches = matches + 1
                    Exit For
                End If
            Next h
        Next c
        If matches >= 20 Then
            headerRow = R
            ValidarEncabezados = True
            Exit Function
        End If
    Next R
    ValidarEncabezados = False
End Function

' Canoniza un encabezado: minúsculas, sin tildes, sin espacios extra
' Acepta "pocentaje" = "porcentaje" como equivalentes
Private Function CanonHeader(ByVal s As String) As String
    Dim res As String
    res = LCase$(Trim$(s))
    ' quitar tildes
    res = Replace(res, Chr(225), "a"): res = Replace(res, Chr(233), "e")
    res = Replace(res, Chr(237), "i"): res = Replace(res, Chr(243), "o")
    res = Replace(res, Chr(250), "u")
    res = Replace(res, Chr(193), "a"): res = Replace(res, Chr(201), "e")
    res = Replace(res, Chr(205), "i"): res = Replace(res, Chr(211), "o")
    res = Replace(res, Chr(218), "u")
    ' typo del sistema: acepta "pocentaje" como "porcentaje"
    res = Replace(res, "pocentaje", "porcentaje")
    CanonHeader = res
End Function

' ============================================================
'  POWER QUERY
'  Transforma tipos: fechas como date, montos/cantidad/precio
'  como number, todo lo demás como text.
'  Normaliza tildes en encabezados y acepta ambas formas del typo.
' ============================================================
Private Function M_Operaciones_PQ() As String
    Dim m As String
    m = "let" & vbCrLf
    m = m & "    Origen = Excel.CurrentWorkbook(){[Name=""Operaciones_Raw""]}[Content]," & vbCrLf

    ' Normalizar encabezados: quitar tildes + corregir typo
    m = m & "    #""Encabezados"" =" & vbCrLf
    m = m & "        Table.TransformColumnNames(Origen, each" & vbCrLf
    m = m & "            let t0  = Text.Trim(_)," & vbCrLf
    m = m & "                t1  = Text.Replace(t0,""" & Chr(225) & """,""a"")," & vbCrLf
    m = m & "                t2  = Text.Replace(t1,""" & Chr(233) & """,""e"")," & vbCrLf
    m = m & "                t3  = Text.Replace(t2,""" & Chr(237) & """,""i"")," & vbCrLf
    m = m & "                t4  = Text.Replace(t3,""" & Chr(243) & """,""o"")," & vbCrLf
    m = m & "                t5  = Text.Replace(t4,""" & Chr(250) & """,""u"")," & vbCrLf
    m = m & "                t6  = Text.Replace(t5,""" & Chr(193) & """,""A"")," & vbCrLf
    m = m & "                t7  = Text.Replace(t6,""" & Chr(201) & """,""E"")," & vbCrLf
    m = m & "                t8  = Text.Replace(t7,""" & Chr(205) & """,""I"")," & vbCrLf
    m = m & "                t9  = Text.Replace(t8,""" & Chr(211) & """,""O"")," & vbCrLf
    m = m & "                t10 = Text.Replace(t9,""" & Chr(218) & """,""U"")," & vbCrLf
    m = m & "                t11 = Text.Replace(t10,""Pocentaje"",""Porcentaje"")" & vbCrLf
    m = m & "            in t11)," & vbCrLf

    ' Parsear fechas explícitamente como DD/MM/YYYY
    ' Las fechas llegan como texto "15/01/2026"; si llegaran como date se pasan directo.
    ' Esto evita que el locale en-US invierta día y mes.
    ' ParseFecha: parsea "DD/MM/YYYY" o "DD-MM-YYYY" explícitamente como día/mes/año
    ' para evitar que el locale en-US invierta día y mes.
    m = m & "    ParseFecha = (v as any) as nullable date =>" & vbCrLf
    m = m & "        if v is date then v" & vbCrLf
    m = m & "        else if v is text then" & vbCrLf
    m = m & "            let pts = Text.Split(Text.Replace(v, ""-"", ""/""), ""/"")," & vbCrLf
    m = m & "                d   = try Number.FromText(pts{0}) otherwise null," & vbCrLf
    m = m & "                mo  = try Number.FromText(pts{1}) otherwise null," & vbCrLf
    m = m & "                y   = try Number.FromText(pts{2}) otherwise null" & vbCrLf
    m = m & "            in if d <> null and mo <> null and y <> null" & vbCrLf
    m = m & "               then #date(y, mo, d) else null" & vbCrLf
    m = m & "        else null," & vbCrLf

    m = m & "    #""FechasParsed"" = Table.TransformColumns(#""Encabezados"", {" & vbCrLf
    m = m & "        {""Fecha de Operacion"", each ParseFecha(_), type date}," & vbCrLf
    m = m & "        {""Fecha Liquidacion"",  each ParseFecha(_), type date}," & vbCrLf
    m = m & "        {""Fecha fin Contrato"", each ParseFecha(_), type date}})," & vbCrLf

    ' Tipos del resto de columnas (fechas ya resueltas arriba)
    m = m & "    #""Tipos"" = Table.TransformColumnTypes(#""FechasParsed"", {" & vbCrLf
    m = m & "        {""Portafolio"",                       type text}," & vbCrLf
    m = m & "        {""Codigo de Orden"",                  type text}," & vbCrLf
    m = m & "        {""Codigo ISIN"",                      type text}," & vbCrLf
    m = m & "        {""Codigo SBS"",                       type text}," & vbCrLf
    m = m & "        {""Monto de Operacion Original"",      type number}," & vbCrLf
    m = m & "        {""Monto de Operacion ML"",            type number}," & vbCrLf
    m = m & "        {""Cantidad"",                         type number}," & vbCrLf
    m = m & "        {""Precio"",                           type number}," & vbCrLf
    m = m & "        {""Codigo de Emisor"",                 type text}," & vbCrLf
    m = m & "        {""Operacion"",                        type text}," & vbCrLf
    m = m & "        {""Moneda"",                           type text}," & vbCrLf
    m = m & "        {""Nemonico"",                         type text}," & vbCrLf
    m = m & "        {""Codigo de Tercero"",                type text}," & vbCrLf
    m = m & "        {""Tercero"",                          type text}," & vbCrLf
    m = m & "        {""Monto Nominal Operacion Original"", type number}," & vbCrLf
    m = m & "        {""Monto Nominal Operacion ML"",       type number}," & vbCrLf
    m = m & "        {""Total de Comisiones"",              type number}," & vbCrLf
    m = m & "        {""Plaza"",                            type text}," & vbCrLf
    m = m & "        {""Tipo Tasa"",                        type text}," & vbCrLf
    m = m & "        {""Porcentaje Tasa"",                  type number}" & vbCrLf
    m = m & "    }, ""es-PE"")," & vbCrLf

    ' Filtrar PRECANCELACION TITULOS UNICOS
    m = m & "    #""Filtrado"" = Table.SelectRows(#""Tipos"", each" & vbCrLf
    m = m & "        Text.Upper(Text.Trim([Operacion])) <> ""PRECANCELACION TITULOS UNICOS"")" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "    #""Filtrado"""
    M_Operaciones_PQ = m
End Function

' ============================================================
'  CARGAR CONSULTA PQ EN HOJA
' ============================================================
Private Sub CargarConsultaEnHoja(ByVal queryName As String, _
                                  ByVal sheetName As String, _
                                  ByVal tableName As String)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet, lo As ListObject, qt As QueryTable

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = sheetName
    Else
        Do While ws.ListObjects.Count > 0
            ws.ListObjects(1).Unlist
        Loop
        Dim qtEx As QueryTable
        For Each qtEx In ws.QueryTables: qtEx.Delete: Next qtEx
        ws.Cells.ClearContents
    End If

    Dim connStr As String
    connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;" & _
              "Data Source=$Workbook$;Location=" & queryName & _
              ";Extended Properties="""";"

    On Error Resume Next
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:=connStr, Destination:=ws.Range("A1"))
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "No se pudo crear la tabla para la consulta '" & queryName & "'.", vbCritical
        Exit Sub
    End If

    On Error GoTo RefreshErr
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .PreserveFormatting = True
        .AdjustColumnWidth = False
        .RefreshStyle = xlInsertDeleteCells
        .Refresh BackgroundQuery:=False
    End With
    On Error GoTo 0

    On Error Resume Next
    lo.name = tableName
    lo.TableStyle = "TableStyleLight8"
    On Error GoTo 0
    Exit Sub

RefreshErr:
    MsgBox "No se pudo actualizar la consulta '" & queryName & "'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    On Error GoTo 0
End Sub

Private Sub UpsertQuery(ByVal qName As String, ByVal mCode As String)
    Dim q As WorkbookQuery
    On Error Resume Next
    Set q = ThisWorkbook.Queries(qName)
    On Error GoTo 0
    If q Is Nothing Then
        ThisWorkbook.Queries.Add name:=qName, Formula:=mCode
    Else
        q.Formula = mCode
    End If
End Sub

' ============================================================
'  PREPARAR HOJA OPERACIONES_RAW
' ============================================================
Private Sub PrepararHojaOperaciones(ByRef wsOut As Worksheet)
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Operaciones_Raw")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsOut.name = "Operaciones_Raw"
    Else
        Do While wsOut.ListObjects.Count > 0
            wsOut.ListObjects(1).Unlist
        Loop
        wsOut.Cells.Clear
    End If
End Sub

' ============================================================
'  DETECCIÓN AUTOMÁTICA DE PERÍODO
' ============================================================
Private Sub AutodetectarPeriodo()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsOp As Worksheet

    On Error Resume Next
    Set wsOp = wb.Worksheets("Operaciones")
    On Error GoTo 0
    If wsOp Is Nothing Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsOp.ListObjects("Operaciones")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub

    Dim fechaCol As Long: fechaCol = ColIdx(lo, "Fecha de Operacion")
    If fechaCol = 0 Then Exit Sub

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim db As Range: Set db = lo.DataBodyRange
    Dim i As Long, fv As Variant, y As Long, m As Long

    For i = 1 To db.Rows.Count
        fv = db.Cells(i, fechaCol).Value
        If IsDate(fv) Then
            y = Year(CDate(fv)): m = Month(CDate(fv))
            dict(y & "-" & Format$(m, "00")) = Array(y, m)
        End If
    Next i

    If dict.Count = 0 Then Exit Sub

    Dim anioMax As Long, k As Variant, arrK As Variant
    anioMax = 0
    For Each k In dict.keys
        arrK = dict(k)
        If arrK(0) > anioMax Then anioMax = arrK(0)
    Next k

    Dim mesesDelAnio As New Collection, mesUnico As Long
    For Each k In dict.keys
        Dim arrM As Variant: arrM = dict(k)
        If arrM(0) = anioMax Then mesesDelAnio.Add arrM(1)
    Next k

    If dict.Count > 1 Then
        Dim lista As String, arr As Variant
        Dim keys() As String: ReDim keys(0 To dict.Count - 1)
        i = 0
        For Each k In dict.keys: keys(i) = CStr(k): i = i + 1: Next k
        Dim j As Long, tmp As String
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If keys(j) < keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            Next j
        Next i
        For i = 0 To UBound(keys)
            arr = dict(keys(i))
            lista = lista & "  " & Chr(149) & "  " & NombreMesES(arr(1)) & " " & arr(0) & vbCrLf
        Next i
        MsgBox "El archivo contiene " & dict.Count & " meses distintos:" & vbCrLf & vbCrLf & _
               lista & vbCrLf & "Se actualiz" & Chr(243) & " a modo Anual (" & anioMax & ")." & vbCrLf & _
               "Puede cambiar el filtro manualmente.", vbInformation, "M" & Chr(250) & "ltiples meses"
    End If

    On Error Resume Next
    wb.Names("A" & Chr(241) & "o").RefersToRange.Value = anioMax
    If mesesDelAnio.Count = 1 Then
        mesUnico = mesesDelAnio(1)
        wb.Names("Mes").RefersToRange.Value = NombreMesES(mesUnico)
        wb.Names("TipoInforme").RefersToRange.Value = "Mensual"
        wb.Names("PeriodoActual").RefersToRange.Value = NombreMesES(mesUnico) & " " & anioMax
    Else
        wb.Names("TipoInforme").RefersToRange.Value = "Anual"
        wb.Names("PeriodoActual").RefersToRange.Value = "Anual " & anioMax
    End If
    On Error GoTo 0
End Sub

' ============================================================
'  FORMATO DE FECHAS EN TABLA FINAL
' ============================================================
Private Sub AplicarFormatosFecha(lo As ListObject)
    Dim cols As Variant
    cols = Array("Fecha de Operacion", "Fecha Liquidacion", "Fecha fin Contrato")
    Dim c As Variant, idx As Long
    For Each c In cols
        idx = ColIdx(lo, CStr(c))
        If idx > 0 Then
            If Not lo.ListColumns(idx).DataBodyRange Is Nothing Then
                lo.ListColumns(idx).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"
            End If
        End If
    Next c
End Sub

' ============================================================
'  DEBUG
' ============================================================
Private Sub GenerarHojaDebug(wsDest As Worksheet, ruta As String)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsD As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Debug_Import").Delete
    Application.DisplayAlerts = True
    Set wsD = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsD.name = "Debug_Import"
    On Error GoTo 0
    wsD.Range("A1").Value = "Archivo":  wsD.Range("B1").Value = ruta
    wsD.Range("A2").Value = "Filas Raw": wsD.Range("B2").Value = _
        IIf(wsDest.ListObjects("Operaciones_Raw").DataBodyRange Is Nothing, _
            0, wsDest.ListObjects("Operaciones_Raw").DataBodyRange.Rows.Count)
    wsD.Range("A3").Value = "Cols":     wsD.Range("B3").Value = N_COLS
    wsD.Range("A4").Value = "Headers:"
    wsD.Range("A5").Resize(1, N_COLS).Value = wsDest.ListObjects("Operaciones_Raw").headerRowRange.Value
    wsD.Columns.AutoFit
End Sub

' ============================================================
'  HELPERS PÚBLICOS
' ============================================================

Public Function ColIdx(lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).name, colName, vbTextCompare) = 0 Then
            ColIdx = i: Exit Function
        End If
    Next i
    Dim low As String: low = LCase$(colName)
    For i = 1 To lo.ListColumns.Count
        If InStr(LCase$(lo.ListColumns(i).name), low) > 0 Then
            ColIdx = i: Exit Function
        End If
    Next i
    ColIdx = 0
End Function

Public Function NombreMesES(ByVal m As Long) As String
    Select Case m
        Case 1:  NombreMesES = "Enero"
        Case 2:  NombreMesES = "Febrero"
        Case 3:  NombreMesES = "Marzo"
        Case 4:  NombreMesES = "Abril"
        Case 5:  NombreMesES = "Mayo"
        Case 6:  NombreMesES = "Junio"
        Case 7:  NombreMesES = "Julio"
        Case 8:  NombreMesES = "Agosto"
        Case 9:  NombreMesES = "Septiembre"
        Case 10: NombreMesES = "Octubre"
        Case 11: NombreMesES = "Noviembre"
        Case 12: NombreMesES = "Diciembre"
        Case Else: NombreMesES = ""
    End Select
End Function

Public Function MesNumero(ByVal s As String) As Long
    Select Case UCase$(Left$(Trim$(s) & "   ", 3))
        Case "ENE": MesNumero = 1
        Case "FEB": MesNumero = 2
        Case "MAR": MesNumero = 3
        Case "ABR": MesNumero = 4
        Case "MAY": MesNumero = 5
        Case "JUN": MesNumero = 6
        Case "JUL": MesNumero = 7
        Case "AGO": MesNumero = 8
        Case "SEP", "SET": MesNumero = 9
        Case "OCT": MesNumero = 10
        Case "NOV": MesNumero = 11
        Case "DIC": MesNumero = 12
        Case Else: MesNumero = 0
    End Select
End Function

Private Function EsTextoOHtml(wb As Workbook) As Boolean
    Select Case wb.FileFormat
        Case xlHtml, xlCurrentPlatformText, xlTextMac, xlTextWindows, xlTextMSDOS, _
             xlCSV, xlCSVWindows, xlCSVMac, xlCSVMSDOS, xlCSVUTF8, xlUnicodeText
            EsTextoOHtml = True
        Case Else
            EsTextoOHtml = False
    End Select
End Function

Private Function PickFilePath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccionar archivo de Operaciones SAF (.XLS, .XLSX)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb"
        If .Show <> -1 Then Exit Function
        PickFilePath = .SelectedItems(1)
    End With
End Function