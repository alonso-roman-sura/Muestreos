' ========== modUniversoPorTipo.bas ==========
Option Explicit

' ============================================================
'  Calcula UniversoPN, UniversoPJ, TamañoPob, TamañoMuestraPN
'  y TamañoMuestraPJ contando todos los registros de la tabla
'  Contratos sin filtro de período, ya que el archivo importado
'  corresponde siempre al período de análisis completo.
'  PN = Tipo empieza por "N" (Natural)
'  PJ = Tipo empieza por "J" (Jurídico)
' ============================================================
Public Sub TamañoPoblacion()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsC As Worksheet
    Dim lo As ListObject
    Dim tipoCol As Long
    Dim db As Range
    Dim i As Long, total As Long, contN As Long, contJ As Long
    Dim tipoVal As String, initial As String

    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Set wsC = wb.Worksheets("Contratos")
    On Error GoTo 0
    If wsC Is Nothing Then GoTo Cleanup

    On Error Resume Next
    Set lo = wsC.ListObjects("Contratos")
    On Error GoTo 0
    If lo Is Nothing Then GoTo Cleanup
    If lo.DataBodyRange Is Nothing Then GoTo Cleanup

    tipoCol = ColIdx(lo, "Tipo")
    If tipoCol = 0 Then
        MsgBox "No se encontr" & Chr(243) & " la columna 'Tipo' en la tabla 'Contratos'.", vbCritical
        GoTo Cleanup
    End If

    ' --- Contar sin filtro de período ---
    Set db = lo.DataBodyRange
    total = 0: contN = 0: contJ = 0

    For i = 1 To db.Rows.Count
        tipoVal = Trim$(CStr(db.Cells(i, tipoCol).Value))
        If Len(tipoVal) > 0 Then
            total = total + 1
            initial = UCase$(Left$(tipoVal, 1))
            If initial = "N" Then contN = contN + 1
            If initial = "J" Then contJ = contJ + 1
        End If
    Next i

    ' --- Guardar universos ---
    On Error Resume Next
    wb.Names("Tama" & Chr(241) & "oPob").RefersToRange.Value = total
    wb.Names("UniversoPN").RefersToRange.Value = contN
    wb.Names("UniversoPJ").RefersToRange.Value = contJ

    ' --- Leer parámetros Cochran ---
    Dim Z As Double, pVal As Double, E As Double
    Z = Val(CStr(wb.Names("Z").RefersToRange.Value))
    pVal = Val(CStr(wb.Names("p").RefersToRange.Value))
    E = Val(CStr(wb.Names("E").RefersToRange.Value))
    On Error GoTo 0

    If Z = 0 Then Z = 1.96
    If pVal = 0 Then pVal = 0.5
    If E = 0 Then E = 0.29

    ' --- Calcular tamaños de muestra ---
    Dim nmPN As String: nmPN = "Tama" & Chr(241) & "oMuestraPN"
    Dim nmPJ As String: nmPJ = "Tama" & Chr(241) & "oMuestraPJ"

    On Error Resume Next
    wb.Names(nmPN).RefersToRange.Value = CochranN(contN, Z, pVal, E)
    wb.Names(nmPJ).RefersToRange.Value = CochranN(contJ, Z, pVal, E)
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

Private Function CochranN(ByVal N As Long, ByVal Z As Double, _
                            ByVal p As Double, ByVal E As Double) As Long
    If N <= 0 Or E <= 0 Or Z <= 0 Then Exit Function
    Dim num As Double: num = N * Z ^ 2 * p * (1 - p)
    Dim den As Double: den = (N - 1) * E ^ 2 + Z ^ 2 * p * (1 - p)
    If den = 0 Then Exit Function
    CochranN = CLng(Application.WorksheetFunction.RoundUp(num / den, 0))
End Function
