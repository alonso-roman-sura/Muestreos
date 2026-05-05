' ========== modImportarDatos.bas ==========
Option Explicit

Private Const DEBUG_IMPORT As Boolean = False

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

    ' -- Abrir workbook ----------------------------------------
    paso = "Abrir workbook"
    Set wbOrigen = Workbooks.Open(Filename:=ruta, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)

    ' -- Buscar hoja válida ------------------------------------
    paso = "Buscar hoja de datos"
    Dim wsOrigen As Worksheet
    Set wsOrigen = BuscarHojaRescates(wbOrigen)

    If wsOrigen Is Nothing Then
        wbOrigen.Close False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "No se encontr" & Chr(243) & " ninguna hoja v" & Chr(225) & "lida en el archivo." & vbCrLf & vbCrLf & _
               "Se busc" & Chr(243) & " una hoja llamada 'RESCATES' y como alternativa" & vbCrLf & _
               "cualquier hoja visible con las columnas esperadas (TIPOPERSONA, CUC, FONDO...)." & vbCrLf & vbCrLf & _
               "Verifique que el archivo sea un reporte de Rescates SAF v" & Chr(225) & "lido.", _
               vbCritical, "Hoja no encontrada"
        Exit Sub
    End If

    ' -- Localizar bloque de datos -----------------------------
    paso = "Encontrar datos en hoja '" & wsOrigen.name & "'"
    Dim startCell As Range
    Dim rngDatos As Range
    Set rngDatos = EncontrarDatos(wsOrigen, startCell)

    If rngDatos Is Nothing Then
        wbOrigen.Close False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "La hoja '" & wsOrigen.name & "' fue encontrada pero no contiene datos." & vbCrLf & vbCrLf & _
               "Las cabeceras existen pero no hay filas de datos debajo de ellas." & vbCrLf & vbCrLf & _
               "Verifique que el archivo tenga registros de rescates cargados.", _
               vbCritical, "Sin datos"
        Exit Sub
    End If

    If rngDatos.Rows.Count < 2 Then
        wbOrigen.Close False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "El archivo no contiene filas de datos (solo cabeceras)." & vbCrLf & _
               "Verifique que sea el archivo correcto.", vbCritical, "Sin datos"
        Exit Sub
    End If

    ' -- Preparar hoja destino ---------------------------------
    paso = "Preparar hoja Rescates"
    Dim wsDest As Worksheet
    PrepararHojaRescates wsDest

    paso = "Copiar valores (" & rngDatos.Rows.Count - 1 & " filas x " & rngDatos.Columns.Count & " cols)"
    wsDest.Range("A1").Resize(rngDatos.Rows.Count, rngDatos.Columns.Count).Value = rngDatos.Value

    ' -- Cerrar origen -----------------------------------------
    paso = "Cerrar workbook origen"
    wbOrigen.Close False
    Set wbOrigen = Nothing

    ' -- Crear ListObject --------------------------------------
    paso = "Crear ListObject"
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsDest.ListObjects("Rescates")
    If Not lo Is Nothing Then lo.Unlist: Set lo = Nothing
    On Error GoTo ERR_HANDLER
    Set lo = wsDest.ListObjects.Add(xlSrcRange, wsDest.Range("A1").CurrentRegion, , xlYes)
    lo.name = "Rescates"

    If lo.DataBodyRange Is Nothing Then
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "Los datos se copiaron pero la tabla qued" & Chr(243) & " vac" & Chr(237) & "a." & vbCrLf & _
               "Revise el archivo origen.", vbCritical, "Tabla vac" & Chr(237) & "a"
        Exit Sub
    End If

    ' -- Formato fechas y autofit ------------------------------
    paso = "Formato fechas y autofit"
    Dim cFO As Long: cFO = BuscarFechaOperacionEnLO(lo)
    If cFO > 0 Then lo.ListColumns(cFO).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"
    Dim fNames As Variant
    fNames = Array("FECHA PROCESO", "FECHA DE PAGO", "FECHA SOLICITUD", "FECHA CREACION")
    Dim fn As Variant, cFN As Long
    For Each fn In fNames
        cFN = ColIdx(lo, CStr(fn))
        If cFN > 0 Then
            If Not lo.ListColumns(cFN).DataBodyRange Is Nothing Then
                lo.ListColumns(cFN).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"
            End If
        End If
    Next fn
    lo.Range.Columns.AutoFit

    If DEBUG_IMPORT Then GenerarHojaDebug wsDest, lo, ruta

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' -- Autodetectar período ----------------------------------
    paso = "Autodetectar per" & Chr(237) & "odo"
    Dim continuar As Boolean
    continuar = AutodetectarPeriodo(lo)
    If Not continuar Then Exit Sub

    ' -- Calcular universo ------------------------------------
    paso = "Tama" & Chr(241) & "oPoblacion"
    TamañoPoblacion

    Dim filas As Long: filas = lo.DataBodyRange.Rows.Count
    MsgBox "Rescates SAF importados correctamente." & vbCrLf & vbCrLf & _
           "Hoja: " & wsDest.name & vbCrLf & _
           "Registros: " & filas & vbCrLf & _
           "Per" & Chr(237) & "odo: " & ObtenerPeriodoActual(), _
           vbInformation, "Importaci" & Chr(243) & "n completada"
    Exit Sub

ERR_HANDLER:
    Dim errDesc As String: errDesc = "[" & paso & "] " & Err.Description
    On Error Resume Next
    If Not wbOrigen Is Nothing Then wbOrigen.Close False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error al importar los datos:" & vbCrLf & vbCrLf & errDesc, _
           vbCritical, "Error de importaci" & Chr(243) & "n"
End Sub

' ============================================================
'  BUSCAR HOJA VÁLIDA
'  1) Hoja llamada "RESCATES" (case-insensitive) con estructura
'  2) Fallback: primera hoja visible con estructura válida
' ============================================================
Private Function BuscarHojaRescates(wb As Workbook) As Worksheet
    Dim sh As Worksheet

    For Each sh In wb.Worksheets
        If sh.Visible = xlSheetVisible Then
            If UCase$(sh.name) = "RESCATES" Then
                If TieneEstructuraRescates(sh) Then
                    Set BuscarHojaRescates = sh
                    Exit Function
                End If
            End If
        End If
    Next sh

    For Each sh In wb.Worksheets
        If sh.Visible = xlSheetVisible Then
            If TieneEstructuraRescates(sh) Then
                Set BuscarHojaRescates = sh
                Exit Function
            End If
        End If
    Next sh

    Set BuscarHojaRescates = Nothing
End Function

Private Function TieneEstructuraRescates(ws As Worksheet) As Boolean
    Dim R As Long
    For R = 1 To Application.Min(30, ws.Rows.Count)
        If TieneCabecerasValidas(ws.Cells(R, 1).Resize(1, 50)) Then
            TieneEstructuraRescates = True
            Exit Function
        End If
    Next R
    TieneEstructuraRescates = False
End Function

' ============================================================
'  LOCALIZAR BLOQUE DE DATOS
'  Escanea hasta fila 30 buscando cabeceras.
'  Usa End(xlUp) en múltiples columnas para la última fila real.
'  Devuelve Nothing si no hay cabeceras o no hay datos.
' ============================================================
Private Function EncontrarDatos(ws As Worksheet, ByRef startCell As Range) As Range
    Dim R As Long

    For R = 1 To Application.Min(30, ws.Rows.Count)
        If TieneCabecerasValidas(ws.Cells(R, 1).Resize(1, 50)) Then

            Dim lastCol As Long
            lastCol = ws.Cells(R, ws.Columns.Count).End(xlToLeft).Column
            If lastCol < 5 Then GoTo SiguienteFila

            ' Buscar primera columna no vacía del header para evitar col A fantasma
            Dim firstCol As Long: firstCol = 1
            Dim fc As Long
            For fc = 1 To Application.Min(lastCol, 20)
                If Len(Trim$(CStr(ws.Cells(R, fc).Value))) > 0 Then
                    firstCol = fc: Exit For
                End If
            Next fc

            ' Calcular última fila real usando End(xlUp) en múltiples columnas
            Dim lastRow As Long: lastRow = R
            Dim tryCol As Long, tryRow As Long
            For tryCol = firstCol To Application.Min(lastCol, firstCol + 10)
                tryRow = ws.Cells(ws.Rows.Count, tryCol).End(xlUp).Row
                If tryRow > lastRow Then lastRow = tryRow
            Next tryCol

            If lastRow <= R Then
                Set EncontrarDatos = Nothing
                Exit Function
            End If

            Set startCell = ws.Cells(R, firstCol)
            Set EncontrarDatos = ws.Range(ws.Cells(R, firstCol), ws.Cells(lastRow, lastCol))
            Exit Function
        End If
SiguienteFila:
    Next R

    Set EncontrarDatos = Nothing
End Function

Private Function TieneCabecerasValidas(headerRow As Range) As Boolean
    Dim esperadas As Variant
    esperadas = Array("TIPOPERSONA", "FONDO", "CUC", "MONTO", "CUOTAS", _
                      "PROMOTOR", "CUENTA", "ESTADO", "TRASPASO")
    Dim maxCol As Long: maxCol = Application.Min(50, headerRow.Columns.Count)
    Dim rng As Range
    Set rng = headerRow.Cells(1, 1).Resize(1, maxCol)
    Dim e As Variant
    For Each e In esperadas
        If BuscarColPorNombre(rng, CStr(e)) > 0 Then
            TieneCabecerasValidas = True
            Exit Function
        End If
    Next e
    TieneCabecerasValidas = False
End Function

' ============================================================
'  PREPARAR HOJA RESCATES
' ============================================================
Private Sub PrepararHojaRescates(ByRef wsOut As Worksheet)
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Rescates")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsOut.name = "Rescates"
    Else
        Do While wsOut.ListObjects.Count > 0
            wsOut.ListObjects(1).Unlist
        Loop
        wsOut.Cells.Clear
    End If
End Sub

' ============================================================
'  DETECCIÓN AUTOMÁTICA DE PERÍODO
'  Un mes   ? escribe etiqueta en PeriodoActual, retorna True
'  Múltiples? pregunta continuar
'             Sí  ? escribe rango, retorna True
'             No  ? limpia datos importados, retorna False
' ============================================================
Private Function AutodetectarPeriodo(lo As ListObject) As Boolean
    AutodetectarPeriodo = True
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim fechaCol As Long
    fechaCol = BuscarFechaOperacionEnLO(lo)
    If fechaCol = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim db As Range: Set db = lo.DataBodyRange
    Dim i As Long, fv As Variant, y As Long, m As Long

    For i = 1 To db.Rows.Count
        fv = db.Cells(i, fechaCol).Value
        Dim dFv As Date
        Dim fechaOK As Boolean: fechaOK = False
        On Error Resume Next
        If IsDate(fv) Then
            dFv = CDate(fv): fechaOK = (Err.Number = 0)
        ElseIf IsNumeric(fv) Then
            Dim dbl As Double: dbl = CDbl(fv)
            If dbl > 1 And dbl < 2958466 Then
                dFv = CDate(dbl): fechaOK = (Err.Number = 0)
            End If
        End If
        On Error GoTo 0
        If fechaOK Then
            y = Year(dFv): m = Month(dFv)
            dict(y & "-" & Format$(m, "00")) = Array(y, m)
        End If
    Next i

    If dict.Count = 0 Then Exit Function

    Dim keys() As String: ReDim keys(0 To dict.Count - 1)
    i = 0
    Dim k As Variant
    For Each k In dict.keys: keys(i) = CStr(k): i = i + 1: Next k
    Dim j As Long, tmp As String
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i

    If dict.Count > 1 Then
        Dim discontinuo As Boolean: discontinuo = False
        For i = 1 To UBound(keys)
            Dim arrPrev As Variant: arrPrev = dict(keys(i - 1))
            Dim arrCurr As Variant: arrCurr = dict(keys(i))
            Dim yEsp As Long: yEsp = arrPrev(0)
            Dim mEsp As Long: mEsp = arrPrev(1) + 1
            If mEsp > 12 Then mEsp = 1: yEsp = yEsp + 1
            If arrCurr(0) <> yEsp Or arrCurr(1) <> mEsp Then
                discontinuo = True: Exit For
            End If
        Next i

        Dim lista As String, arr As Variant
        For i = 0 To UBound(keys)
            arr = dict(keys(i))
            lista = lista & "  " & Chr(149) & "  " & NombreMesES(arr(1)) & " " & arr(0) & vbCrLf
        Next i

        Dim msgExtra As String: msgExtra = ""
        If discontinuo Then
            msgExtra = vbCrLf & "Los meses no son consecutivos. Esto puede indicar un " & _
                       "archivo incorrecto o datos faltantes." & vbCrLf
        End If

        Dim resp As VbMsgBoxResult
        resp = MsgBox( _
            "El archivo contiene " & dict.Count & " meses distintos:" & vbCrLf & vbCrLf & _
            lista & msgExtra & vbCrLf & _
            "Se recomienda importar archivos de un solo mes." & vbCrLf & vbCrLf & _
            Chr(191) & "Desea continuar de todas formas?", _
            vbYesNo + vbExclamation + vbDefaultButton2, _
            "Archivo con m" & Chr(250) & "ltiples meses")
        If resp <> vbYes Then
            On Error Resume Next
            lo.Unlist
            ThisWorkbook.Worksheets("Rescates").Cells.Clear
            On Error GoTo 0
            MsgBox "Importaci" & Chr(243) & "n cancelada. No se cargaron datos.", _
                   vbInformation, "Cancelado"
            AutodetectarPeriodo = False
            Exit Function
        End If
    End If

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
End Function

Private Function NombreExiste(ByVal nm As String) As Boolean
    Dim n As name
    On Error Resume Next
    Set n = ThisWorkbook.Names(nm)
    NombreExiste = Not (n Is Nothing)
    On Error GoTo 0
End Function

Private Function ObtenerPeriodoActual() As String
    On Error Resume Next
    ObtenerPeriodoActual = CStr(ThisWorkbook.Names("PeriodoActual").RefersToRange.Value)
    On Error GoTo 0
End Function

' ============================================================
'  DEBUG
' ============================================================
Private Sub GenerarHojaDebug(wsDest As Worksheet, lo As ListObject, ruta As String)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsD As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Debug_Import").Delete
    Application.DisplayAlerts = True
    Set wsD = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsD.name = "Debug_Import"
    On Error GoTo 0
    wsD.Range("A1").Value = "Archivo":           wsD.Range("B1").Value = ruta
    wsD.Range("A2").Value = "Filas importadas":  wsD.Range("B2").Value = _
        IIf(lo.DataBodyRange Is Nothing, 0, lo.DataBodyRange.Rows.Count)
    wsD.Range("A3").Value = "Columnas":          wsD.Range("B3").Value = lo.ListColumns.Count
    wsD.Range("A4").Value = "Encabezados:"
    wsD.Range("A5").Resize(1, lo.ListColumns.Count).Value = lo.headerRowRange.Value
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

Public Function BuscarFechaOperacionEnLO(lo As ListObject) As Long
    ' Búsqueda con Canon: elimina tildes y espacios antes de comparar
    ' "FECHA OPERACIÓN" y "FECHA OPERACION" quedan ambas como "FECHAOPERACION"
    Dim targets As Variant
    targets = Array("FECHAPROCESO", "FECHAOPERACION", "FECHADEOPERACION", "FECHAOPER")
    Dim i As Long, hNorm As String
    For i = 1 To lo.ListColumns.Count
        hNorm = Canon(lo.ListColumns(i).name)
        Dim t As Variant
        For Each t In targets
            If hNorm = CStr(t) Or InStr(hNorm, CStr(t)) > 0 Then
                BuscarFechaOperacionEnLO = i: Exit Function
            End If
        Next t
    Next i
    ' Fallback: primera columna con "fecha" en el nombre
    For i = 1 To lo.ListColumns.Count
        If InStr(LCase$(lo.ListColumns(i).name), "fecha") > 0 Then
            BuscarFechaOperacionEnLO = i: Exit Function
        End If
    Next i
    BuscarFechaOperacionEnLO = 0
End Function

Public Function BuscarColPorNombre(headerRow As Range, ParamArray nombres() As Variant) As Long
    Dim c As Range, n As Variant, hNorm As String, nNorm As String
    For Each c In headerRow.Cells
        hNorm = Canon(CStr(c.Value))
        For Each n In nombres
            nNorm = Canon(CStr(n))
            If hNorm = nNorm Or (Len(nNorm) > 0 And InStr(hNorm, nNorm) > 0) Then
                BuscarColPorNombre = c.Column - headerRow.Cells(1, 1).Column + 1
                Exit Function
            End If
        Next n
    Next c
    BuscarColPorNombre = 0
End Function

Public Function Canon(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, Chr(160), "")  ' espacio de no separación
    s = Replace(s, Chr(193), "A"): s = Replace(s, Chr(201), "E")
    s = Replace(s, Chr(205), "I"): s = Replace(s, Chr(211), "O")
    s = Replace(s, Chr(218), "U"): s = Replace(s, Chr(209), "N")
    s = Replace(s, " ", ""): s = Replace(s, "_", "")
    s = Replace(s, "-", ""): s = Replace(s, "/", "")
    s = Replace(s, ".", "")
    Canon = s
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

Private Function PickFilePath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccionar archivo de Rescates SAF (.XLS, .XLSX)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb"
        If .Show <> -1 Then Exit Function
        PickFilePath = .SelectedItems(1)
    End With
End Function