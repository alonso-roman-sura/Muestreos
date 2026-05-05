' ========== modUniversoPorTipo.bas ==========
Option Explicit

' ============================================================
'  Calcula el universo de Operaciones SAF filtrando por período
'  y excluyendo "PRECANCELACION TITULOS UNICOS".
'  No hay segmentación PN/PJ: un solo universo.
' ============================================================
Public Sub TamañoPoblacion()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsOp As Worksheet

    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Set wsOp = wb.Worksheets("Operaciones")
    On Error GoTo 0
    If wsOp Is Nothing Then GoTo Cleanup

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsOp.ListObjects("Operaciones")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then GoTo Cleanup

    Dim opCol As Long
    opCol = BuscarColExacta(lo, "Operacion")
    If opCol = 0 Then
        MsgBox "No se encontr" & Chr(243) & " la columna 'Operacion'.", vbCritical
        GoTo Cleanup
    End If

    ' Contar todo excluyendo PRECANCELACION TITULOS UNICOS
    ' Sin filtro de período: el archivo importado ya es del período correcto
    Dim db As Range: Set db = lo.DataBodyRange
    Dim total As Long: total = 0
    Dim i As Long, opVal As String

    For i = 1 To db.Rows.Count
        opVal = UCase$(Trim$(CStr(db.Cells(i, opCol).Value)))
        If opVal <> "PRECANCELACION TITULOS UNICOS" Then
            total = total + 1
        End If
    Next i

    On Error Resume Next
    wb.Names("Universo").RefersToRange.Value = total

    Dim Z As Double, pVal As Double, E As Double
    Z = val(CStr(wb.Names("Z").RefersToRange.Value))
    pVal = val(CStr(wb.Names("p").RefersToRange.Value))
    E = val(CStr(wb.Names("E").RefersToRange.Value))
    On Error GoTo 0

    If Z = 0 Then Z = 1.96
    If pVal = 0 Then pVal = 0.5
    If E = 0 Then E = 0.29

    Dim nmTam As String: nmTam = "Tama" & Chr(241) & "oMuestra"
    On Error Resume Next
    wb.Names(nmTam).RefersToRange.Value = CochranN(total, Z, pVal, E)
    On Error GoTo 0

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error en Tama" & Chr(241) & "oPoblacion: " & Err.Number & " - " & Err.Description, vbCritical
End Sub
' ============================================================
'  HELPERS
' ============================================================

Private Function CochranN(ByVal n As Long, ByVal Z As Double, _
                            ByVal p As Double, ByVal E As Double) As Long
    If n <= 0 Or E <= 0 Or Z <= 0 Then Exit Function
    Dim num As Double: num = n * Z ^ 2 * p * (1 - p)
    Dim den As Double: den = (n - 1) * E ^ 2 + Z ^ 2 * p * (1 - p)
    If den = 0 Then Exit Function
    CochranN = CLng(Application.WorksheetFunction.RoundUp(num / den, 0))
End Function


' Busca coincidencia exacta únicamente (sin fallback parcial)
' para evitar que "Operacion" encuentre "Fecha de Operacion"
Private Function BuscarColExacta(lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).name, colName, vbTextCompare) = 0 Then
            BuscarColExacta = i: Exit Function
        End If
    Next i
    BuscarColExacta = 0
End Function