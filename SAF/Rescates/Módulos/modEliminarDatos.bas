' ========== modEliminarDatos.bas ==========
Option Explicit

Private Const HOJAS_PROTEGIDAS As String = "Instrucciones|Muestra"

' ============================================================
'  ENTRADA DEL BOTÓN "Eliminar Datos"
' ============================================================
Public Sub EliminarDatos()

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "Esta acci" & Chr(243) & "n limpiar" & Chr(225) & " el archivo por completo:" & vbCrLf & vbCrLf & _
        "   " & Chr(149) & "  Todas las hojas excepto Instrucciones y Muestra" & vbCrLf & _
        "   " & Chr(149) & "  Los valores generados en la hoja Muestra" & vbCrLf & vbCrLf & _
        Chr(191) & "Desea continuar?", _
        vbYesNo + vbExclamation + vbDefaultButton2, "Eliminar datos")
    If resp <> vbYes Then Exit Sub

    resp = MsgBox( _
        "Esta operaci" & Chr(243) & "n no se puede deshacer." & vbCrLf & vbCrLf & _
        Chr(191) & "Confirma que desea limpiar el archivo?", _
        vbYesNo + vbCritical + vbDefaultButton2, "Confirmar eliminaci" & Chr(243) & "n")
    If resp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo FIN

    EliminarHojas
    LimpiarHojaMuestra

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "El archivo ha sido limpiado y est" & Chr(225) & " listo para recibir nuevos datos.", _
           vbInformation, "Listo"
    Exit Sub

FIN:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If Err.Number <> 0 Then
        MsgBox "Error inesperado durante la limpieza:" & vbCrLf & Err.Description, vbCritical, "Error"
    End If
End Sub

' ============================================================
'  ELIMINAR HOJAS (excepto las protegidas)
' ============================================================
Private Sub EliminarHojas()
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If Not EsHojaProtegida(ThisWorkbook.Worksheets(i).name) Then
            ThisWorkbook.Worksheets(i).Delete
        End If
    Next i
End Sub

Private Function EsHojaProtegida(ByVal nombre As String) As Boolean
    Dim partes() As String: partes = Split(HOJAS_PROTEGIDAS, "|")
    Dim p As Variant
    For Each p In partes
        If LCase$(Trim$(CStr(p))) = LCase$(Trim$(nombre)) Then
            EsHojaProtegida = True: Exit Function
        End If
    Next p
    EsHojaProtegida = False
End Function

' ============================================================
'  LIMPIAR HOJA MUESTRA
'  Vacía: TamañoPob, UniversoPN, UniversoPJ,
'          TamañoMuestraPN, TamañoMuestraPJ, PeriodoActual.
'  Conserva: Z, p, E, Mes, Año, TipoInforme (controles del usuario).
'  Borra (Clear): grillas Muestra1_PN, Muestra1_PJ.
' ============================================================
Private Sub LimpiarHojaMuestra()
    Dim wsM As Worksheet
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets("Muestra")
    On Error GoTo 0
    If wsM Is Nothing Then Exit Sub

    LimpiarCelda "Tama" & Chr(241) & "oPob"
    LimpiarCelda "UniversoPN"
    LimpiarCelda "UniversoPJ"
    LimpiarCelda "Tama" & Chr(241) & "oMuestraPN"
    LimpiarCelda "Tama" & Chr(241) & "oMuestraPJ"
    LimpiarCelda "PeriodoActual"

    LimpiarGrilla "Muestra1_PN", 5
    LimpiarGrilla "Muestra1_PJ", 5
End Sub

Private Sub LimpiarCelda(ByVal nmNombre As String)
    Dim nm As name
    On Error Resume Next
    Set nm = ThisWorkbook.Names(nmNombre)
    On Error GoTo 0
    If nm Is Nothing Then Exit Sub
    Dim cel As Range
    On Error Resume Next
    Set cel = nm.RefersToRange
    On Error GoTo 0
    If Not cel Is Nothing Then cel.ClearContents
End Sub

Private Sub LimpiarGrilla(ByVal nmNombre As String, ByVal nCols As Long)
    Dim nm As name
    On Error Resume Next
    Set nm = ThisWorkbook.Names(nmNombre)
    On Error GoTo 0
    If nm Is Nothing Then Exit Sub

    Dim inicio As Range
    On Error Resume Next
    Set inicio = nm.RefersToRange
    On Error GoTo 0
    If inicio Is Nothing Then Exit Sub

    Dim ws As Worksheet: Set ws = inicio.Parent
    Dim lastRow As Long, lr As Long, c As Long
    lastRow = inicio.Row
    For c = 0 To nCols - 1
        lr = ws.Cells(ws.Rows.Count, inicio.Column + c).End(xlUp).Row
        If lr > lastRow Then lastRow = lr
    Next c
    If lastRow < inicio.Row Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(inicio, ws.Cells(lastRow, inicio.Column + nCols - 1))
    On Error Resume Next
    rng.UnMerge
    On Error GoTo 0
    rng.Clear
End Sub