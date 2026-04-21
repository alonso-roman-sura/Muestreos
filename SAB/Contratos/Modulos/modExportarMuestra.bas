' ========== modExportarMuestra.bas ==========
Option Explicit

' ============================================================
'  ENTRADA DEL BOTÓN "Generar Tabla con las Muestras"
'
'  Flujo:
'  1. Lee los números aleatorios desde Muestra1_PN / Muestra1_PJ.
'  2. Usa esos números como posiciones dentro del universo filtrado
'     por tipo (N o J), sin filtro de período (el archivo ya es
'     del período correcto).
'  3. Exporta las filas seleccionadas a hojas separadas.
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
        MsgBox "No se han generado los n" & Chr(250) & "meros de muestra." & vbCrLf & _
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
        MsgBox "No se encontr" & Chr(243) & " la tabla 'Contratos' o est" & Chr(225) & " vac" & Chr(237) & "a.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo FIN

    Dim cntPN As Long, cntPJ As Long
    cntPN = ExportarTipo(wb, lo, "N", "Muestra_Contratos_PN", celdaPN)
    cntPJ = ExportarTipo(wb, lo, "J", "Muestra_Contratos_PJ", celdaPJ)

    MsgBox "Exportaci" & Chr(243) & "n completada." & vbCrLf & _
           "PN: " & cntPN & " fila(s)." & vbCrLf & _
           "PJ: " & cntPJ & " fila(s).", vbInformation

FIN:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ============================================================
'  Exporta filas para un tipo (N o J) usando los números
'  aleatorios como índices dentro del subuniverso por tipo.
'  No filtra por período: todos los registros del tipo dado
'  son elegibles, ya que el archivo importado es del período
'  correcto.
' ============================================================
Private Function ExportarTipo(wb As Workbook, lo As ListObject, _
                               ByVal inicial As String, _
                               ByVal hojaDestino As String, _
                               celdaInicio As Range) As Long
    Dim tipoCol As Long
    tipoCol = ColIdx(lo, "Tipo")
    If tipoCol = 0 Then Exit Function

    ' Construir índices del subuniverso filtrado solo por tipo
    Dim db As Range: Set db = lo.DataBodyRange
    Dim universoIdx() As Long
    ReDim universoIdx(1 To db.Rows.Count)
    Dim N As Long: N = 0
    Dim i As Long, tipoVal As String

    For i = 1 To db.Rows.Count
        tipoVal = Trim$(CStr(db.Cells(i, tipoCol).Value))
        If UCase$(Left$(tipoVal, 1)) = UCase$(inicial) Then
            N = N + 1
            universoIdx(N) = i
        End If
    Next i

    If N = 0 Then Exit Function

    ' Leer números de muestra desde la grilla (5 columnas hacia abajo)
    Dim nums() As Long
    nums = LeerNumerosGrilla(celdaInicio, 5)
    If UBound(nums) = 0 Then Exit Function

    ' Mapear número ? índice real en la tabla
    Dim selIdx() As Long, k As Long, c As Long
    ReDim selIdx(1 To UBound(nums))
    k = 0
    For c = 1 To UBound(nums)
        If nums(c) >= 1 And nums(c) <= N Then
            k = k + 1
            selIdx(k) = universoIdx(nums(c))
        End If
    Next c
    If k = 0 Then Exit Function
    ReDim Preserve selIdx(1 To k)

    ' Crear/limpiar hoja destino
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
    lo.HeaderRowRange.Copy
    wsDest.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    ' Filas seleccionadas
    Dim dstRow As Long: dstRow = 2
    For c = 1 To k
        db.Rows(selIdx(c)).Copy
        wsDest.Cells(dstRow, 1).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        dstRow = dstRow + 1
    Next c

    ' Crear tabla y aplicar estilo
    Dim loT As ListObject
    Set loT = wsDest.ListObjects.Add(xlSrcRange, wsDest.Range("A1").CurrentRegion, , xlYes)
    loT.name = hojaDestino
    On Error Resume Next
    loT.TableStyle = "TableStyleLight9"
    On Error GoTo 0
    loT.Range.Columns.AutoFit

    ' Formato de fecha
    Dim cF As Long: cF = ColIdx(loT, "Fecha de Ingreso")
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
'  HELPERS
' ============================================================

Private Function ColIdx(lo As ListObject, ByVal colName As String) As Long
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