' ========== mod2.bas  (ExportarMuestra) ==========
Option Explicit

' ============================================================
'  ENTRADA DEL BOTÓN "Generar Tabla con las Muestras"
'
'  Flujo:
'  1. Construye el universo filtrado por período y tipo (PN/PJ),
'     ordenado igual que la tabla cargada (Fecha ASC, Transac ASC).
'  2. Lee los números aleatorios desde Muestra1_PN / Muestra1_PJ.
'  3. Exporta las filas que corresponden a esos números (posiciones
'     dentro del universo filtrado) a hojas separadas.
' ============================================================
Public Sub ExportarMuestra()
    Dim wb As Workbook: Set wb = ThisWorkbook

    ' Validar que se hayan generado los números de muestra
    Dim celdaPN As Range, celdaPJ As Range
    On Error Resume Next
    Set celdaPN = wb.Names("Muestra1_PN").RefersToRange
    Set celdaPJ = wb.Names("Muestra1_PJ").RefersToRange
    On Error GoTo 0

    If celdaPN Is Nothing Or celdaPJ Is Nothing Then
        MsgBox "No se encontraron los nombres definidos 'Muestra1_PN' / 'Muestra1_PJ'.", vbCritical
        Exit Sub
    End If
    If IsEmpty(celdaPN.Value) Or Len(Trim$(CStr(celdaPN.Value))) = 0 Then
        MsgBox "No se han generado los números de muestra." & vbCrLf & _
               "Primero ejecute 'Seleccionar Muestras'.", vbExclamation, "Sin muestra"
        Exit Sub
    End If

    ' Obtener tabla
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = wb.Worksheets("Contratos")
    On Error GoTo 0
    If wsC Is Nothing Then
        MsgBox "No existe la hoja 'Contratos'.", vbCritical: Exit Sub
    End If
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsC.ListObjects("Contratos")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        MsgBox "No se encontró la tabla 'Contratos' o está vacía.", vbCritical: Exit Sub
    End If

    ' Leer filtros del período
    Dim tipoInforme As String, mesTexto As String
    Dim anioFiltro As Long, mesFiltro As Long
    On Error Resume Next
    tipoInforme = UCase$(Trim$(CStr(wb.Names("TipoInforme").RefersToRange.Value)))
    anioFiltro  = CLng(wb.Names("Año").RefersToRange.Value)
    If tipoInforme = "MENSUAL" Then
        mesTexto  = Trim$(CStr(wb.Names("Mes").RefersToRange.Value))
        mesFiltro = MesNumero(mesTexto)
    Else
        mesFiltro = 0
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents  = False

    On Error GoTo FIN

    Dim cntPN As Long, cntPJ As Long
    cntPN = ExportarTipo(wb, lo, "N", "Muestra_Contratos_PN", celdaPN, anioFiltro, mesFiltro)
    cntPJ = ExportarTipo(wb, lo, "J", "Muestra_Contratos_PJ", celdaPJ, anioFiltro, mesFiltro)

    MsgBox "Exportación completada." & vbCrLf & _
           "PN: " & cntPN & " fila(s)." & vbCrLf & _
           "PJ: " & cntPJ & " fila(s).", vbInformation

FIN:
    Application.EnableEvents  = True
    Application.ScreenUpdating = True
End Sub

' ============================================================
'  Exporta filas para un tipo (N o J) usando los números
'  aleatorios como índices dentro del universo filtrado.
'  Devuelve la cantidad de filas exportadas.
' ============================================================
Private Function ExportarTipo(wb As Workbook, lo As ListObject, _
                               ByVal inicial As String, _
                               ByVal hojaDestino As String, _
                               celdaInicio As Range, _
                               ByVal anioFiltro As Long, _
                               ByVal mesFiltro As Long) As Long
    Dim fechaCol As Long, tipoCol As Long
    fechaCol = ColIdx(lo, "Fecha")
    tipoCol  = ColIdx(lo, "Tipo Persona")
    If fechaCol = 0 Or tipoCol = 0 Then Exit Function

    ' 1) Construir índices de filas del universo filtrado (mismo orden que la tabla)
    Dim db As Range: Set db = lo.DataBodyRange
    Dim universoIdx() As Long
    ReDim universoIdx(1 To db.Rows.Count)
    Dim n As Long: n = 0
    Dim i As Long, fechaVal As Variant, tipoVal As String

    For i = 1 To db.Rows.Count
        fechaVal = db.Cells(i, fechaCol).Value
        If IsDate(fechaVal) Then
            If Year(fechaVal) = anioFiltro Then
                If mesFiltro = 0 Or Month(fechaVal) = mesFiltro Then
                    tipoVal = Trim$(CStr(db.Cells(i, tipoCol).Value))
                    If UCase$(Left$(tipoVal, 1)) = UCase$(inicial) Then
                        n = n + 1
                        universoIdx(n) = i
                    End If
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function

    ' 2) Leer números de muestra desde la grilla (5 columnas, hacia abajo)
    Dim nums() As Long
    nums = LeerNumerosGrilla(celdaInicio, 5)
    If UBound(nums) = 0 Then Exit Function

    ' 3) Mapear número → índice real en la tabla
    Dim selIdx() As Long, k As Long, c As Long
    ReDim selIdx(1 To UBound(nums))
    k = 0
    For c = 1 To UBound(nums)
        If nums(c) >= 1 And nums(c) <= n Then
            k = k + 1
            selIdx(k) = universoIdx(nums(c))
        End If
    Next c
    If k = 0 Then Exit Function
    ReDim Preserve selIdx(1 To k)

    ' 4) Crear/limpiar hoja destino
    Dim wsDest As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsDest = wb.Worksheets(hojaDestino)
    If Not wsDest Is Nothing Then wsDest.Delete
    Set wsDest = Nothing
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsDest = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsDest.Name = hojaDestino

    ' 5) Encabezados
    lo.HeaderRowRange.Copy
    wsDest.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    ' 6) Filas seleccionadas
    Dim dstRow As Long: dstRow = 2
    For c = 1 To k
        db.Rows(selIdx(c)).Copy
        wsDest.Cells(dstRow, 1).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        dstRow = dstRow + 1
    Next c

    ' 7) Crear tabla y autofit
    Dim loT As ListObject
    Set loT = wsDest.ListObjects.Add(xlSrcRange, wsDest.Range("A1").CurrentRegion, , xlYes)
    loT.Name = hojaDestino
    On Error Resume Next
    loT.TableStyle = "TableStyleLight9"
    On Error GoTo 0
    loT.Range.Columns.AutoFit

    ' Columna Fecha: formatear como fecha
    Dim cF As Long: cF = ColIdx(loT, "Fecha")
    If cF > 0 Then loT.ListColumns(cF).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"

    ExportarTipo = k
End Function

' ============================================================
'  Lee números de una grilla de nCols columnas hacia abajo.
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

' ============================================================
'  HELPERS (duplicados de mod3 para que el módulo sea autónomo)
' ============================================================

Private Function ColIdx(lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, colName, vbTextCompare) = 0 Then
            ColIdx = i: Exit Function
        End If
    Next i
    Dim low As String: low = LCase$(colName)
    For i = 1 To lo.ListColumns.Count
        If InStr(LCase$(lo.ListColumns(i).Name), low) > 0 Then
            ColIdx = i: Exit Function
        End If
    Next i
    ColIdx = 0
End Function

Private Function MesNumero(ByVal s As String) As Long
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
        Case Else:  MesNumero = 0
    End Select
End Function
