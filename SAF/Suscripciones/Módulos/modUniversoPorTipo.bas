' ========== modUniversoPorTipo_Suscripciones.bas ==========
Option Explicit

Public Sub TamañoPoblacion()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsC As Worksheet

    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Set wsC = wb.Worksheets("Suscripciones")
    On Error GoTo 0
    If wsC Is Nothing Then GoTo Cleanup

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsC.ListObjects("Suscripciones")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then GoTo Cleanup

    Dim tipoCol As Long
    tipoCol = ColIdx(lo, "TIPO PERSONA")
    If tipoCol = 0 Then tipoCol = ColIdx(lo, "TIPOPERSONA")
    If tipoCol = 0 Then
        MsgBox "No se encontr" & Chr(243) & " la columna 'TIPO PERSONA'.", vbCritical
        GoTo Cleanup
    End If

    ' Sin filtro de período: contar todo el archivo importado
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

    On Error Resume Next
    wb.Names("Tama" & Chr(241) & "oPob").RefersToRange.Value = total
    wb.Names("UniversoPN").RefersToRange.Value = contN
    wb.Names("UniversoPJ").RefersToRange.Value = contJ

    Dim Z As Double, pVal As Double, E As Double
    Z = val(CStr(wb.Names("Z").RefersToRange.Value))
    pVal = val(CStr(wb.Names("p").RefersToRange.Value))
    E = val(CStr(wb.Names("E").RefersToRange.Value))
    On Error GoTo 0

    If Z = 0 Then Z = 1.96
    If pVal = 0 Then pVal = 0.5
    If E = 0 Then E = 0.29

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

' Fórmula de Cochran con corrección de población finita
Private Function CochranN(ByVal n As Long, ByVal Z As Double, _
                           ByVal p As Double, ByVal E As Double) As Long
    If n <= 0 Or E <= 0 Then CochranN = 0: Exit Function
    Dim n0 As Double
    n0 = (Z ^ 2 * p * (1 - p)) / (E ^ 2)
    If n = 0 Then CochranN = 0: Exit Function
    Dim nCorr As Double
    nCorr = n0 / (1 + (n0 - 1) / n)
    CochranN = CLng(Application.WorksheetFunction.RoundUp(nCorr, 0))
    If CochranN > n Then CochranN = n
End Function

Public Function NormalizarTipoPersona(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, Chr(160), "")
    If s = "NAT" Or InStr(s, "NATURAL") > 0 Then NormalizarTipoPersona = "N": Exit Function
    If s = "MAN" Or InStr(s, "MANCOMUNADO") > 0 Then NormalizarTipoPersona = "N": Exit Function
    If s = "JUR" Or InStr(s, "JURIDIC") > 0 Then NormalizarTipoPersona = "J": Exit Function
    If s = "N" Or s = "M" Then NormalizarTipoPersona = "N": Exit Function
    If s = "J" Then NormalizarTipoPersona = "J": Exit Function
    NormalizarTipoPersona = ""
End Function

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