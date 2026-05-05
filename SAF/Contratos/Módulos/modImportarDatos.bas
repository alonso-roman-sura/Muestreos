' ========== modImportarDatos.bas ==========
Option Explicit

Private Const COL_CODIGO     As String = "CODIGO"
Private Const COL_NOMBRE     As String = "NOMBRE DEL PARTICIPE"
Private Const COL_TIPDOC     As String = "TIPO DOCUMENTO"
Private Const COL_NUMDOC     As String = "NUMERO DOCUMENTO"
Private Const COL_TIPPERSONA As String = "TIPO PERSONA"
Private Const COL_TIPCLIENTE As String = "TIPO DE CLIENTE"
Private Const COL_AGENCIA    As String = "AGENCIA"
Private Const COL_CODPROM    As String = "CODIGO_PROMOTOR"
Private Const COL_PROMOTOR   As String = "PROMOTOR"
Private Const COL_NUEVO      As String = "NUEVO/CLIENTE"
Private Const COL_FECHA      As String = "FECHA_APERTURA_FONDO"
Private Const COL_CONTRATOS  As String = "CONTRATOS_PREVIOS"

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
    Set wbOrigen = Workbooks.Open(Filename:=ruta, ReadOnly:=True)

    paso = "Leer hoja origen"
    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Sheets(1)

    paso = "Encontrar datos"
    Dim rngDatos As Range
    Set rngDatos = EncontrarDatos(wsOrigen)
    If rngDatos Is Nothing Then
        wbOrigen.Close False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "No se encontraron las cabeceras esperadas.", vbCritical, "Formato no reconocido"
        Exit Sub
    End If

    paso = "Preparar hoja Contratos"
    Dim wsDest As Worksheet
    PrepararHojaContratos wsDest

    ' Localizar columna NUMERO DOCUMENTO en el bloque de datos origen
    paso = "Localizar columna NUMERO DOCUMENTO"
    Dim numDocCol As Long: numDocCol = 0
    Dim hdr As Range
    For Each hdr In rngDatos.Rows(1).Cells
        If Canon(CStr(hdr.Value)) = Canon(COL_NUMDOC) Then
            numDocCol = hdr.Column - rngDatos.Columns(1).Column + 1
            Exit For
        End If
    Next hdr

    ' Formatear esa columna como texto ANTES de pegar
    ' para que los ceros iniciales no sean eliminados por Excel
    If numDocCol > 0 Then
        wsDest.Columns(numDocCol).NumberFormat = "@"
    End If

    paso = "Copiar valores (" & rngDatos.Rows.Count & " filas x " & rngDatos.Columns.Count & " cols)"
    wsDest.Range("A1").Resize(rngDatos.Rows.Count, rngDatos.Columns.Count).Value = rngDatos.Value

    ' Re-leer NUMERO DOCUMENTO como .Text desde el origen para preservar
    ' exactamente lo que muestra Excel (ej. "01234567" con cero inicial).
    ' Si el archivo fuente lo guardó como número ya perdió el cero;
    ' en ese caso .Text devuelve el número sin el cero, que es lo mejor que se puede hacer.
    If numDocCol > 0 Then
        Dim srcCol As Range: Set srcCol = rngDatos.Columns(numDocCol)
        Dim r As Long
        For r = 2 To rngDatos.Rows.Count
            Dim srcTxt As String: srcTxt = srcCol.Cells(r, 1).Text
            If Len(Trim$(srcTxt)) > 0 Then
                wsDest.Cells(r, numDocCol).Value = srcTxt
            End If
        Next r
    End If

    paso = "Cerrar workbook origen"
    wbOrigen.Close False
    Set wbOrigen = Nothing

    paso = "Crear ListObject"
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsDest.ListObjects("Contratos")
    If Not lo Is Nothing Then lo.Unlist: Set lo = Nothing
    On Error GoTo ERR_HANDLER
    Set lo = wsDest.ListObjects.Add(xlSrcRange, _
             wsDest.Range("A1").CurrentRegion, , xlYes)
    lo.name = "Contratos"

    paso = "Formato fecha y autofit"
    Dim cF As Long: cF = ColIdx(lo, COL_FECHA)
    If cF > 0 Then lo.ListColumns(cF).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"
    lo.Range.Columns.AutoFit

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    paso = "Autodetectar per" & Chr(237) & "odo"
    AutodetectarPeriodo lo

    paso = "Tama" & Chr(241) & "oPoblacion"
    TamañoPoblacion

    MsgBox "Datos SAF cargados correctamente.", vbInformation
    Exit Sub

ERR_HANDLER:
    Dim errDesc As String: errDesc = "[" & paso & "] " & Err.Description
    On Error Resume Next
    If Not wbOrigen Is Nothing Then wbOrigen.Close False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error al cargar los datos:" & vbCrLf & vbCrLf & errDesc, _
           vbCritical, "Error de importaci" & Chr(243) & "n"
End Sub

' ============================================================
'  Localiza el bloque de datos en la hoja origen.
'  Busca cabeceras en B11 primero; si no coinciden, intenta A1.
'  Devuelve Nothing si no encuentra cabeceras válidas.
' ============================================================
Private Function EncontrarDatos(ws As Worksheet) As Range
    Dim anclas As Variant
    anclas = Array("B11", "A1")

    Dim a As Variant, cel As Range, rng As Range
    For Each a In anclas
        On Error Resume Next
        Set cel = ws.Range(CStr(a))
        On Error GoTo 0
        If cel Is Nothing Then GoTo SiguienteAncla

        Set rng = cel.CurrentRegion
        If rng Is Nothing Then GoTo SiguienteAncla
        If rng.Rows.Count < 2 Or rng.Columns.Count < 4 Then GoTo SiguienteAncla

        If TieneCabecerasValidas(rng.Rows(1)) Then
            Set EncontrarDatos = rng
            Exit Function
        End If
SiguienteAncla:
    Next a
    Set EncontrarDatos = Nothing
End Function

' Comprueba que al menos 2 de las cabeceras esperadas estén en la fila
Private Function TieneCabecerasValidas(headerRow As Range) As Boolean
    Dim esperadas As Variant
    esperadas = Array(COL_CODIGO, COL_TIPPERSONA, COL_FECHA, COL_NOMBRE, COL_NUMDOC)
    Dim encontradas As Long, e As Variant
    For Each e In esperadas
        If FindHeaderCol(headerRow, CStr(e)) > 0 Then
            encontradas = encontradas + 1
            If encontradas >= 2 Then
                TieneCabecerasValidas = True
                Exit Function
            End If
        End If
    Next e
    TieneCabecerasValidas = False
End Function

Private Function FindHeaderCol(headerRow As Range, ByVal colName As String) As Long
    Dim c As Range
    For Each c In headerRow.Cells
        If Canon(CStr(c.Value)) = Canon(colName) Then
            FindHeaderCol = c.Column - headerRow.Cells(1, 1).Column + 1
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

' Normaliza texto: mayúsculas, sin acentos, sin separadores
Private Function Canon(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, Chr(193), "A"): s = Replace(s, Chr(201), "E")
    s = Replace(s, Chr(205), "I"): s = Replace(s, Chr(211), "O")
    s = Replace(s, Chr(218), "U"): s = Replace(s, Chr(209), "N")
    s = Replace(s, " ", ""): s = Replace(s, "_", "")
    s = Replace(s, "-", ""): s = Replace(s, "/", "")
    Canon = s
End Function

' ============================================================
'  Prepara la hoja Contratos
' ============================================================
Private Sub PrepararHojaContratos(ByRef wsOut As Worksheet)
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Contratos")
    On Error GoTo 0

    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsOut.name = "Contratos"
    Else
        Do While wsOut.ListObjects.Count > 0
            wsOut.ListObjects(1).Unlist
        Loop
        wsOut.Cells.Clear
    End If
End Sub

Private Function PickFilePath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccionar archivo SAF (.XLS, .XLSX)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb"
        If .Show <> -1 Then Exit Function
        PickFilePath = .SelectedItems(1)
    End With
End Function

' ============================================================
'  DETECCIÓN DE PERÍODO
' ============================================================
Private Sub AutodetectarPeriodo(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim fechaCol As Long
    fechaCol = ColIdx(lo, COL_FECHA)
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

    If dict.Count > 1 Then
        Dim lista As String, k As Variant, arr As Variant
        For Each k In dict.keys
            arr = dict(k)
            lista = lista & "  " & Chr(149) & "  " & NombreMesES(arr(1)) & " " & arr(0) & vbCrLf
        Next k
        Dim resp As VbMsgBoxResult
        resp = MsgBox( _
            "El archivo contiene " & dict.Count & " meses distintos:" & vbCrLf & vbCrLf & _
            lista & vbCrLf & _
            "Se recomienda importar archivos de un solo mes." & vbCrLf & vbCrLf & _
            Chr(191) & "Desea continuar de todas formas?", _
            vbYesNo + vbExclamation + vbDefaultButton2, _
            "Archivo con m" & Chr(250) & "ltiples meses")
        If resp <> vbYes Then
            MsgBox "Importaci" & Chr(243) & "n cancelada. Los datos no fueron cargados.", _
                   vbInformation, "Cancelado"
            On Error Resume Next
            Dim wsClear As Worksheet
            Set wsClear = ThisWorkbook.Worksheets("Contratos")
            If Not wsClear Is Nothing Then
                Do While wsClear.ListObjects.Count > 0
                    wsClear.ListObjects(1).Unlist
                Loop
                wsClear.Cells.Clear
            End If
            On Error GoTo 0
            Exit Sub
        End If
    End If

    Dim keys() As String
    ReDim keys(0 To dict.Count - 1)
    i = 0
    For Each k In dict.keys
        keys(i) = CStr(k): i = i + 1
    Next k
    Dim j As Long, tmp As String
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i

    Dim etiqueta As String
    If dict.Count = 1 Then
        Dim a1 As Variant: a1 = dict(keys(0))
        etiqueta = NombreMesES(a1(1)) & " " & a1(0)
    Else
        Dim aF As Variant: aF = dict(keys(0))
        Dim aL As Variant: aL = dict(keys(UBound(keys)))
        etiqueta = NombreMesES(aF(1)) & " " & aF(0) & " - " & NombreMesES(aL(1)) & " " & aL(0)
    End If

    On Error Resume Next
    wb.Names("PeriodoActual").RefersToRange.Value = etiqueta
    On Error GoTo 0
End Sub

' ============================================================
'  HELPERS públicos (accesibles desde otros módulos)
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