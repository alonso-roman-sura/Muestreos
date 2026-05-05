' ========== modExportarMuestra.bas ==========
Option Explicit

' ============================================================
'  ENTRADA DEL BOTÓN "Generar Tabla con las Muestras"
'
'  Flujo:
'  1. Lee los números aleatorios desde Muestra1_PN / Muestra1_PJ.
'  2. Construye el subuniverso por tipo (PN = NAT+MAN, PJ = JUR)
'     en el mismo orden que la tabla importada.
'  3. Exporta las filas seleccionadas a hojas separadas.
' ============================================================
Public Sub ExportarMuestra()
    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim celdaPN As Range, celdaPJ As Range
    On Error Resume Next
    Set celdaPN = wb.Names("Muestra1_PN").RefersToRange
    Set celdaPJ = wb.Names("Muestra1_PJ").RefersToRange
    On Error GoTo 0

    If celdaPN Is Nothing Or celdaPJ Is Nothing Then
        MsgBox "No se encontraron los nombres 'Muestra1_PN' / 'Muestra1_PJ'.", vbCritical
        Exit Sub
    End If
    If IsEmpty(celdaPN.Value) Or Len(Trim$(CStr(celdaPN.Value))) = 0 Then
        MsgBox "No se han generado los n" & Chr(250) & "meros de muestra." & vbCrLf & _
               "Primero ejecute 'Seleccionar Muestras'.", vbExclamation, "Sin muestra"
        Exit Sub
    End If

    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = wb.Worksheets("Contratos")
    On Error GoTo 0
    If wsC Is Nothing Then
        MsgBox "No existe la hoja 'Contratos'. Importe los datos primero.", vbCritical
        Exit Sub
    End If

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsC.ListObjects("Contratos")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        MsgBox "La tabla 'Contratos' est" & Chr(225) & " vac" & Chr(237) & "a." & vbCrLf & _
               "Importe los datos primero.", vbCritical, "Sin datos"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo FIN

    Dim cntPN As Long, cntPJ As Long
    Dim sufijo As String: sufijo = SufijoHoja()
    cntPN = ExportarTipo(wb, lo, "N", "Muestra_Contratos_SAF_PN" & sufijo, celdaPN)
    cntPJ = ExportarTipo(wb, lo, "J", "Muestra_Contratos_SAF_PJ" & sufijo, celdaPJ)

    If cntPN > 0 Or cntPJ > 0 Then
        MsgBox "Exportaci" & Chr(243) & "n completada." & vbCrLf & _
               "PN (NAT+MAN): " & cntPN & " fila(s)." & vbCrLf & _
               "PJ (JUR): " & cntPJ & " fila(s).", vbInformation
    End If

FIN:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error al exportar la muestra:" & vbCrLf & Err.Description, vbCritical, "Error"
    End If
End Sub

' ============================================================
'  Exporta filas para un tipo usando los números aleatorios
'  como índices dentro del subuniverso filtrado por tipo.
' ============================================================
Private Function ExportarTipo(wb As Workbook, lo As ListObject, _
                               ByVal tipoCod As String, _
                               ByVal hojaDestino As String, _
                               celdaInicio As Range) As Long
    Dim tipoCol As Long
    tipoCol = ColIdxLocal(lo, "TIPO PERSONA")
    If tipoCol = 0 Then
        MsgBox "No se encontr" & Chr(243) & " la columna 'TIPO PERSONA' en la tabla Contratos.", _
               vbCritical, "Error": Exit Function
    End If

    Dim db As Range: Set db = lo.DataBodyRange
    Dim nCols As Long: nCols = lo.ListColumns.Count

    Dim universoIdx() As Long
    ReDim universoIdx(1 To db.Rows.Count)
    Dim n As Long: n = 0
    Dim i As Long, tipoVal As String

    For i = 1 To db.Rows.Count
        tipoVal = Trim$(UCase$(CStr(db.Cells(i, tipoCol).Value)))
        If NormalizarTipoPersona(tipoVal) = tipoCod Then
            n = n + 1
            universoIdx(n) = i
        End If
    Next i

    If n = 0 Then
        MsgBox "No hay registros de tipo '" & tipoCod & "' en la tabla Contratos." & vbCrLf & _
               "Verifique que los datos est" & Chr(233) & "n cargados correctamente.", _
               vbExclamation, "Universo vac" & Chr(237) & "o": Exit Function
    End If

    Dim nums() As Long
    nums = LeerNumerosGrilla(celdaInicio, 5)
    If UBound(nums) = 0 Then
        MsgBox "No se encontraron n" & Chr(250) & "meros en la grilla de muestra " & tipoCod & "." & vbCrLf & _
               "Primero ejecute 'Seleccionar Muestras'.", _
               vbExclamation, "Grilla vac" & Chr(237) & "a": Exit Function
    End If

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
        MsgBox "Los n" & Chr(250) & "meros de la muestra " & tipoCod & " est" & Chr(225) & "n fuera del rango del universo (" & n & " registros)." & vbCrLf & _
               "Regenere la muestra con 'Seleccionar Muestras'.", _
               vbExclamation, "N" & Chr(250) & "meros fuera de rango": Exit Function
    End If
    ReDim Preserve selIdx(1 To k)
    ReDim Preserve selPos(1 To k)

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
    wsDest.Cells(1, nCols + 1).Value = "N" & Chr(186) & " en universo " & tipoCod

    ' Formatear columna NUMERO DOCUMENTO como texto antes de escribir
    Dim numDocCol As Long: numDocCol = ColIdxLocal(lo, "NUMERO DOCUMENTO")
    If numDocCol > 0 Then
        wsDest.Columns(numDocCol).NumberFormat = "@"
    End If

    ' Filas seleccionadas
    Dim dstRow As Long: dstRow = 2
    For c = 1 To k
        wsDest.Cells(dstRow, 1).Resize(1, nCols).Value = db.Rows(selIdx(c)).Value

        If numDocCol > 0 Then
            Dim srcTxt As String
            srcTxt = db.Cells(selIdx(c), numDocCol).Text
            If Len(Trim$(srcTxt)) > 0 Then
                wsDest.Cells(dstRow, numDocCol).Value = srcTxt
            End If
        End If

        wsDest.Cells(dstRow, nCols + 1).Value = selPos(c)
        dstRow = dstRow + 1
    Next c

    Dim loT As ListObject
    Set loT = wsDest.ListObjects.Add(xlSrcRange, _
              wsDest.Range("A1").Resize(k + 1, nCols + 1), , xlYes)
    loT.name = hojaDestino
    On Error Resume Next
    loT.TableStyle = IIf(tipoCod = "N", "TableStyleMedium7", "TableStyleMedium3")
    On Error GoTo 0

    Dim cF As Long: cF = ColIdxLocal(loT, "FECHA_APERTURA_FONDO")
    If cF > 0 Then loT.ListColumns(cF).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"

    loT.Range.Columns.AutoFit

    ExportarTipo = k
End Function
' ============================================================
'  HELPERS
' ============================================================

Private Function LeerNumerosGrilla(startCell As Range, ByVal nCols As Long) As Long()
    Dim nums() As Long, cap As Long
    cap = 0
    ReDim nums(0 To 0)
    Dim r As Long, c As Long, v As Variant, filaVacia As Boolean
    r = 0
    Do
        filaVacia = True
        For c = 0 To nCols - 1
            v = startCell.Offset(r, c).Value
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
        r = r + 1
    Loop
    LeerNumerosGrilla = nums
End Function

' Normaliza tipo persona SAF: NAT/MAN ? "N", JUR ? "J"
Private Function NormalizarTipoPersona(ByVal s As String) As String
    s = UCase$(Trim$(s))
    If s = "NAT" Or InStr(s, "NATURAL") > 0 Then NormalizarTipoPersona = "N": Exit Function
    If s = "MAN" Or InStr(s, "MANCOMUNADO") > 0 Then NormalizarTipoPersona = "N": Exit Function
    If s = "JUR" Or InStr(s, "JURIDIC") > 0 Then NormalizarTipoPersona = "J": Exit Function
    If s = "N" Or s = "M" Then NormalizarTipoPersona = "N": Exit Function
    If s = "J" Then NormalizarTipoPersona = "J": Exit Function
    NormalizarTipoPersona = ""
End Function

Private Function ColIdxLocal(lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).name, colName, vbTextCompare) = 0 Then
            ColIdxLocal = i: Exit Function
        End If
    Next i
    Dim low As String: low = LCase$(colName)
    For i = 1 To lo.ListColumns.Count
        If InStr(LCase$(lo.ListColumns(i).name), low) > 0 Then
            ColIdxLocal = i: Exit Function
        End If
    Next i
    ColIdxLocal = 0
End Function

Private Function SufijoHoja() As String
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim periodo As String
    On Error Resume Next
    periodo = Trim$(CStr(wb.Names("PeriodoActual").RefersToRange.Value))
    On Error GoTo 0
    If Len(periodo) = 0 Then Exit Function

    ' Si contiene " - " es multi-mes: no agregar sufijo
    If InStr(periodo, " - ") > 0 Then Exit Function

    ' Formato esperado: "Enero 2026" ? "_Ene26"
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