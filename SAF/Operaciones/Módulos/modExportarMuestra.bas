' ========== modExportarMuestra.bas ==========
Option Explicit

' ============================================================
'  ENTRADA DEL BOTÓN "Generar Tabla con la Muestra"
' ============================================================
Public Sub ExportarMuestra()
    Dim wb As Workbook: Set wb = ThisWorkbook

    ' Validar números de muestra
    Dim celda As Range
    On Error Resume Next
    Set celda = wb.Names("Muestra1").RefersToRange
    On Error GoTo 0
    If celda Is Nothing Then
        MsgBox "No se encontr" & Chr(243) & " el nombre definido 'Muestra1'.", vbCritical
        Exit Sub
    End If
    If IsEmpty(celda.Value) Or Len(Trim$(CStr(celda.Value))) = 0 Then
        MsgBox "No se han generado los n" & Chr(250) & "meros de muestra." & vbCrLf & _
               "Primero ejecute 'Seleccionar Muestra'.", vbExclamation, "Sin muestra"
        Exit Sub
    End If

    ' Validar tabla origen
    Dim wsOp As Worksheet
    On Error Resume Next
    Set wsOp = wb.Worksheets("Operaciones")
    On Error GoTo 0
    If wsOp Is Nothing Then
        MsgBox "No existe la hoja 'Operaciones'. Importe los datos primero.", vbCritical
        Exit Sub
    End If

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsOp.ListObjects("Operaciones")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        MsgBox "La tabla 'Operaciones' est" & Chr(225) & " vac" & Chr(237) & "a." & vbCrLf & _
               "Importe los datos primero.", vbCritical, "Sin datos"
        Exit Sub
    End If

    ' Leer filtros
    Dim tipoInforme As String, anioFiltro As Long, mesFiltro As Long
    On Error Resume Next
    tipoInforme = UCase$(Trim$(CStr(wb.Names("TipoInforme").RefersToRange.Value)))
    anioFiltro = CLng(wb.Names("A" & Chr(241) & "o").RefersToRange.Value)
    If tipoInforme = "MENSUAL" Then
        mesFiltro = MesNumero(Trim$(CStr(wb.Names("Mes").RefersToRange.Value)))
    Else
        mesFiltro = 0
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo FIN

    Dim sufijo As String: sufijo = SufijoHoja()
    Dim cnt As Long
    cnt = ExportarOperaciones(wb, lo, "Muestra_Operaciones_SAF" & sufijo, celda)

    If cnt > 0 Then
        MsgBox "Exportaci" & Chr(243) & "n completada." & vbCrLf & cnt & " operaci" & Chr(243) & "n(es).", vbInformation
    End If

FIN:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error al exportar la muestra:" & vbCrLf & Err.Description, vbCritical, "Error"
    End If
End Sub

' ============================================================
'  Exporta las filas seleccionadas por los números de la grilla.
'  El subuniverso se construye filtrando por período y
'  excluyendo PRECANCELACION TITULOS UNICOS, en el mismo orden
'  que la tabla Operaciones.
' ============================================================
Private Function ExportarOperaciones(wb As Workbook, lo As ListObject, _
                                      ByVal hojaDestino As String, _
                                      celdaInicio As Range) As Long
    Dim opCol As Long
    opCol = BuscarColExacta(lo, "Operacion")
    If opCol = 0 Then
        MsgBox "No se encontr" & Chr(243) & " la columna 'Operacion'.", vbCritical, "Error"
        Exit Function
    End If

    Dim db As Range: Set db = lo.DataBodyRange
    Dim nCols As Long: nCols = lo.ListColumns.Count

    ' Subuniverso: todo excepto PRECANCELACION TITULOS UNICOS
    Dim universoIdx() As Long
    ReDim universoIdx(1 To db.Rows.Count)
    Dim n As Long: n = 0
    Dim i As Long, opVal As String

    For i = 1 To db.Rows.Count
        opVal = UCase$(Trim$(CStr(db.Cells(i, opCol).Value)))
        If opVal <> "PRECANCELACION TITULOS UNICOS" Then
            n = n + 1
            universoIdx(n) = i
        End If
    Next i

    If n = 0 Then
        MsgBox "No hay operaciones para el per" & Chr(237) & "odo seleccionado." & vbCrLf & _
               "Verifique los filtros Mes/A" & Chr(241) & "o/TipoInforme.", _
               vbExclamation, "Universo vac" & Chr(237) & "o"
        Exit Function
    End If

    ' Leer números de la grilla
    Dim nums() As Long
    nums = LeerNumerosGrilla(celdaInicio, 5)
    If UBound(nums) = 0 Then
        MsgBox "No se encontraron n" & Chr(250) & "meros en la grilla de muestra." & vbCrLf & _
               "Primero ejecute 'Seleccionar Muestra'.", vbExclamation, "Grilla vac" & Chr(237) & "a"
        Exit Function
    End If

    ' Mapear número ? fila real
    Dim selIdx() As Long, selPos() As Long
    Dim k As Long, c As Long
    ReDim selIdx(1 To UBound(nums))
    ReDim selPos(1 To UBound(nums))
    k = 0
    For c = 1 To UBound(nums)
        If nums(c) >= 1 And nums(c) <= n Then
            k = k + 1
            selIdx(k) = universoIdx(nums(c))
            selPos(k) = nums(c)
        End If
    Next c
    If k = 0 Then
        MsgBox "Los n" & Chr(250) & "meros de la muestra est" & Chr(225) & "n fuera del rango del universo (" & n & " operaciones)." & vbCrLf & _
               "Regenere la muestra con 'Seleccionar Muestra'.", _
               vbExclamation, "N" & Chr(250) & "meros fuera de rango"
        Exit Function
    End If
    ReDim Preserve selIdx(1 To k)
    ReDim Preserve selPos(1 To k)

    ' Crear hoja destino
    Dim wsDest As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsDest = wb.Worksheets(hojaDestino)
    If Not wsDest Is Nothing Then wsDest.Delete
    Set wsDest = Nothing
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDest = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsDest.name = hojaDestino

    ' Encabezados
    wsDest.Range("A1").Resize(1, nCols).Value = lo.headerRowRange.Value
    wsDest.Cells(1, nCols + 1).Value = "N" & Chr(186) & " en universo"

    ' Filas seleccionadas
    Dim dstRow As Long: dstRow = 2
    For c = 1 To k
        wsDest.Cells(dstRow, 1).Resize(1, nCols).Value = db.Rows(selIdx(c)).Value
        wsDest.Cells(dstRow, nCols + 1).Value = selPos(c)
        dstRow = dstRow + 1
    Next c

    ' Tabla
    Dim loT As ListObject
    Set loT = wsDest.ListObjects.Add(xlSrcRange, _
              wsDest.Range("A1").Resize(k + 1, nCols + 1), , xlYes)
    loT.name = hojaDestino
    On Error Resume Next
    loT.TableStyle = "TableStyleMedium7"
    On Error GoTo 0

    ' Formatos fecha
    Dim cols As Variant
    cols = Array("Fecha de Operacion", "Fecha Liquidacion", "Fecha fin Contrato")
    Dim col As Variant, cF As Long
    For Each col In cols
        cF = ColIdx(loT, CStr(col))
        If cF > 0 Then
            If Not loT.ListColumns(cF).DataBodyRange Is Nothing Then
                loT.ListColumns(cF).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"
            End If
        End If
    Next col

    loT.Range.Columns.AutoFit

    ExportarOperaciones = k
End Function

' ============================================================
'  HELPERS
' ============================================================

Private Function LeerNumerosGrilla(startCell As Range, ByVal nCols As Long) As Long()
    Dim nums() As Long, cap As Long
    cap = 0
    ReDim nums(0 To 0)
    Dim R As Long, c As Long, v As Variant, filaVacia As Boolean
    R = 0
    Do
        filaVacia = True
        For c = 0 To nCols - 1
            v = startCell.Offset(R, c).Value
            If Len(CStr(v)) > 0 Then
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
    LeerNumerosGrilla = nums
End Function

Private Function SufijoHoja() As String
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim periodo As String
    On Error Resume Next
    periodo = Trim$(CStr(wb.Names("PeriodoActual").RefersToRange.Value))
    On Error GoTo 0
    If Len(periodo) = 0 Then Exit Function
    If InStr(periodo, " - ") > 0 Or InStr(periodo, "Anual") > 0 Then Exit Function
    Dim partes() As String: partes = Split(periodo, " ")
    If UBound(partes) < 1 Then Exit Function
    Dim mesAbrev As String: mesAbrev = Left$(partes(0), 3)
    Dim anioAbrev As String
    If Len(partes(1)) >= 4 Then
        anioAbrev = Right$(partes(1), 2)
    Else
        anioAbrev = partes(1)
    End If
    SufijoHoja = "_" & mesAbrev & anioAbrev
End Function


Private Function BuscarColExacta(lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).name, colName, vbTextCompare) = 0 Then
            BuscarColExacta = i: Exit Function
        End If
    Next i
    BuscarColExacta = 0
End Function
