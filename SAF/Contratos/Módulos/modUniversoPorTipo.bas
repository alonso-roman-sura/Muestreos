' ========== modUniversoPorTipo.bas ==========
Option Explicit

' ============================================================
'  Calcula universos y tamaños de muestra SAF.
'  PN = TIPO PERSONA es "NAT" o "MAN" (Mancomunado cuenta como Natural)
'  PJ = TIPO PERSONA es "JUR"
'  Sin filtro de período: el archivo importado ya es del período correcto.
' ============================================================
Public Sub TamañoPoblacion()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsC As Worksheet

    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Set wsC = wb.Worksheets("Contratos")
    On Error GoTo 0
    If wsC Is Nothing Then GoTo Cleanup

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsC.ListObjects("Contratos")
    On Error GoTo 0
    If lo Is Nothing Then GoTo Cleanup
    If lo.DataBodyRange Is Nothing Then GoTo Cleanup

    Dim tipoCol As Long
    tipoCol = ColIdx(lo, "TIPO PERSONA")
    If tipoCol = 0 Then
        MsgBox "No se encontr" & Chr(243) & " la columna 'TIPO PERSONA'.", vbCritical
        GoTo Cleanup
    End If

    ' --- Contar ---
    Dim db As Range: Set db = lo.DataBodyRange
    Dim total As Long, contN As Long, contJ As Long
    Dim i As Long, tipoVal As String, tipoCod As String
    total = 0: contN = 0: contJ = 0

    For i = 1 To db.Rows.Count
        tipoVal = Trim$(UCase$(CStr(db.Cells(i, tipoCol).Value)))
        If Len(tipoVal) > 0 Then
            tipoCod = NormalizarTipoPersona(tipoVal)
            If tipoCod <> "" Then
                total = total + 1
                If tipoCod = "N" Then contN = contN + 1
                If tipoCod = "J" Then contJ = contJ + 1
            End If
        End If
    Next i

    ' --- Guardar universos ---
    On Error Resume Next
    wb.Names("Tama" & Chr(241) & "oPob").RefersToRange.Value = total
    wb.Names("UniversoPN").RefersToRange.Value = contN
    wb.Names("UniversoPJ").RefersToRange.Value = contJ

    ' --- Parámetros Cochran ---
    Dim Z As Double, pVal As Double, e As Double
    Z = val(CStr(wb.Names("Z").RefersToRange.Value))
    pVal = val(CStr(wb.Names("p").RefersToRange.Value))
    e = val(CStr(wb.Names("E").RefersToRange.Value))
    On Error GoTo 0

    If Z = 0 Then Z = 1.96
    If pVal = 0 Then pVal = 0.5
    If e = 0 Then e = 0.29

    Dim nmPN As String: nmPN = "Tama" & Chr(241) & "oMuestraPN"
    Dim nmPJ As String: nmPJ = "Tama" & Chr(241) & "oMuestraPJ"

    On Error Resume Next
    wb.Names(nmPN).RefersToRange.Value = CochranN(contN, Z, pVal, e)
    wb.Names(nmPJ).RefersToRange.Value = CochranN(contJ, Z, pVal, e)
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

' Devuelve "N" (NAT o MAN), "J" (JUR) o "" si no reconoce
Public Function NormalizarTipoPersona(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, Chr(160), "")  ' NBSP

    ' Valores completos del archivo SAF
    If s = "NAT" Or InStr(s, "NATURAL") > 0 Then NormalizarTipoPersona = "N": Exit Function
    If s = "MAN" Or InStr(s, "MANCOMUNADO") > 0 Then NormalizarTipoPersona = "N": Exit Function
    If s = "JUR" Or InStr(s, "JURIDIC") > 0 Or InStr(s, "JUR" & Chr(205) & "DIC") > 0 Then
        NormalizarTipoPersona = "J": Exit Function
    End If

    ' Fallback: iniciales sueltas
    If s = "N" Then NormalizarTipoPersona = "N": Exit Function
    If s = "J" Then NormalizarTipoPersona = "J": Exit Function
    If s = "M" Then NormalizarTipoPersona = "N": Exit Function

    NormalizarTipoPersona = ""
End Function


Private Function CochranN(ByVal n As Long, ByVal Z As Double, _
                            ByVal p As Double, ByVal e As Double) As Long
    If n <= 0 Or e <= 0 Or Z <= 0 Then Exit Function
    Dim num As Double: num = n * Z ^ 2 * p * (1 - p)
    Dim den As Double: den = (n - 1) * e ^ 2 + Z ^ 2 * p * (1 - p)
    If den = 0 Then Exit Function
    CochranN = CLng(Application.WorksheetFunction.RoundUp(num / den, 0))
End Function