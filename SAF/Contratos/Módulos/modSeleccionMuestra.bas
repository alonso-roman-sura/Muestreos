' ========== modSeleccionMuestra.bas ==========
Option Explicit

' ============================================================
'  ENTRADA DEL BOTÓN "Seleccionar Muestras"
' ============================================================
Public Sub SeleccionMuestra()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("¿Está seguro de que desea generar nuevas muestras PN y PJ?", _
                  vbYesNo + vbQuestion, "Confirmar")
    If resp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo FIN

    GenerarMuestra "Muestra1_PN", "Tama" & Chr(241) & "oMuestraPN", "UniversoPN"
    GenerarMuestra "Muestra1_PJ", "Tama" & Chr(241) & "oMuestraPJ", "UniversoPJ"

    MsgBox "Muestras PN y PJ generadas correctamente.", vbInformation

FIN:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error al generar la muestra:" & vbCrLf & Err.Description, vbCritical, "Error"
    End If
End Sub

' ============================================================
'  Genera una muestra para el nombre de inicio dado
' ============================================================
Private Sub GenerarMuestra(ByVal nombreInicio As String, _
                             ByVal nombreTamano As String, _
                             ByVal nombreUniverso As String)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim rngInicio As Range, ws As Worksheet
    Dim tamano As Long, universo As Long

    On Error Resume Next
    Set rngInicio = wb.Names(nombreInicio).RefersToRange
    On Error GoTo 0
    If rngInicio Is Nothing Then
        MsgBox "No existe el nombre definido '" & nombreInicio & "'.", vbCritical
        Exit Sub
    End If
    Set ws = rngInicio.Parent

    On Error Resume Next
    tamano = CLng(wb.Names(nombreTamano).RefersToRange.Value)
    universo = CLng(wb.Names(nombreUniverso).RefersToRange.Value)
    On Error GoTo 0

    If tamano <= 0 Or universo <= 0 Then
        MsgBox "Los valores de '" & nombreTamano & "' y '" & nombreUniverso & "' deben ser > 0.", vbExclamation
        Exit Sub
    End If
    If tamano > universo Then
        MsgBox "'" & nombreTamano & "' no puede ser mayor que '" & nombreUniverso & "'.", vbExclamation
        Exit Sub
    End If

    LimpiarBloque ws, rngInicio, 5

    Dim nums() As Long
    nums = UniqueSortedSample(universo, tamano)
    WriteGrid ws, rngInicio, nums, 5
End Sub

' ============================================================
'  HELPERS
' ============================================================

Private Sub LimpiarBloque(ws As Worksheet, startCell As Range, ByVal nCols As Long)
    Dim c As Long, lastRow As Long, lr As Long
    lastRow = startCell.Row
    For c = 0 To nCols - 1
        lr = ws.Cells(ws.Rows.Count, startCell.Column + c).End(xlUp).Row
        If lr > lastRow Then lastRow = lr
    Next c
    If lastRow < startCell.Row Then Exit Sub
    ws.Range(startCell, ws.Cells(lastRow, startCell.Column + nCols - 1)).Clear
End Sub

Private Function UniqueSortedSample(ByVal universo As Long, ByVal tamano As Long) As Long()
    Dim coll As Collection: Set coll = New Collection
    Dim rnum As Long, i As Long, j As Long, tmp As Long
    Dim v() As Long
    Randomize
    Do While coll.Count < tamano
        rnum = Int(universo * Rnd) + 1
        On Error Resume Next
        coll.Add rnum, CStr(rnum)
        On Error GoTo 0
    Loop
    ReDim v(1 To coll.Count)
    For i = 1 To coll.Count: v(i) = coll(i): Next i
    For i = 1 To UBound(v) - 1
        For j = i + 1 To UBound(v)
            If v(i) > v(j) Then tmp = v(i): v(i) = v(j): v(j) = tmp
        Next j
    Next i
    UniqueSortedSample = v
End Function

' Escribe array en grilla de nCols: borde punteado gris + centrado
Private Sub WriteGrid(ws As Worksheet, startCell As Range, _
                       ByRef v() As Long, ByVal nCols As Long)
    Dim i As Long, r As Long, c As Long, cel As Range, b As Long
    r = startCell.Row: c = startCell.Column
    For i = LBound(v) To UBound(v)
        Set cel = ws.Cells(r, c)
        cel.Value = v(i)
        cel.HorizontalAlignment = xlCenter

        startCell.Copy
        cel.PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        For b = xlEdgeLeft To xlEdgeRight   ' 7 a 10
            With cel.Borders(b)
                .LineStyle = xlDot
                .Weight = xlThin
                .Color = RGB(128, 128, 128)
            End With
        Next b

        c = c + 1
        If c > startCell.Column + (nCols - 1) Then
            c = startCell.Column
            r = r + 1
        End If
    Next i
End Sub