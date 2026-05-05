' ========== modEliminarDatos_Suscripciones.bas ==========
Option Explicit

Public Sub EliminarDatos()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Esta acci" & Chr(243) & "n eliminar" & Chr(225) & " todos los datos importados," & vbCrLf & _
                  "los universos calculados y las muestras generadas." & vbCrLf & vbCrLf & _
                  Chr(191) & "Desea continuar?", _
                  vbYesNo + vbExclamation, "Confirmar eliminaci" & Chr(243) & "n")
    If resp <> vbYes Then Exit Sub

    resp = MsgBox("Confirmaci" & Chr(243) & "n final: se borrar" & Chr(225) & " toda la informaci" & Chr(243) & "n." & vbCrLf & _
                  Chr(191) & "Est" & Chr(225) & " seguro?", _
                  vbYesNo + vbCritical, "Segunda confirmaci" & Chr(243) & "n")
    If resp <> vbYes Then Exit Sub

    Dim wb As Workbook: Set wb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo ErrHandler

    ' 1) Eliminar hoja Suscripciones
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = wb.Worksheets("Suscripciones")
    On Error GoTo ErrHandler
    If Not wsC Is Nothing Then
        wsC.Delete
        Set wsC = Nothing
    End If

    ' 2) Eliminar hojas de muestra exportadas (prefijo conocido)
    '    Recolectar primero, luego borrar (nunca borrar dentro de For Each)
    Dim hojasBorrar() As String
    Dim nHojas As Long: nHojas = 0
    Dim wsLoop As Worksheet
    For Each wsLoop In wb.Worksheets
        If Left$(wsLoop.name, 25) = "Muestra_Suscripciones_PN_" Or _
           Left$(wsLoop.name, 25) = "Muestra_Suscripciones_PJ_" Or _
           wsLoop.name = "Muestra_Suscripciones_PN" Or _
           wsLoop.name = "Muestra_Suscripciones_PJ" Then
            nHojas = nHojas + 1
            ReDim Preserve hojasBorrar(1 To nHojas)
            hojasBorrar(nHojas) = wsLoop.name
        End If
    Next wsLoop

    Dim h As Long
    For h = 1 To nHojas
        On Error Resume Next
        wb.Worksheets(hojasBorrar(h)).Delete
        On Error GoTo ErrHandler
    Next h

    ' 3) Limpiar valores numéricos de nombres definidos
    Dim nmNumericos As Variant
    nmNumericos = Array("UniversoPN", "UniversoPJ", _
                        "Tama" & Chr(241) & "oPob", _
                        "Tama" & Chr(241) & "oMuestraPN", _
                        "Tama" & Chr(241) & "oMuestraPJ")
    Dim nm As Variant
    On Error Resume Next
    For Each nm In nmNumericos
        wb.Names(CStr(nm)).RefersToRange.Value = 0
    Next nm

    ' Limpiar etiqueta de período
    wb.Names("PeriodoActual").RefersToRange.Value = ""
    On Error GoTo ErrHandler

    ' 4) Limpiar grillas de números de muestra (contenido + bordes)
    Dim celdaPN As Range, celdaPJ As Range
    On Error Resume Next
    Set celdaPN = wb.Names("Muestra1_PN").RefersToRange
    Set celdaPJ = wb.Names("Muestra1_PJ").RefersToRange
    On Error GoTo ErrHandler

    If Not celdaPN Is Nothing Then LimpiarGrilla celdaPN
    If Not celdaPJ Is Nothing Then LimpiarGrilla celdaPJ

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Datos eliminados correctamente.", vbInformation
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error al eliminar datos:" & vbCrLf & errNum & " - " & errDesc, vbCritical
End Sub

Private Sub LimpiarGrilla(startCell As Range)
    Dim ws As Worksheet: Set ws = startCell.Parent
    Dim lastRow As Long: lastRow = startCell.Row
    Dim c As Long, R As Long

    ' Calcular última fila usada en las 5 columnas de la grilla
    For c = startCell.Column To startCell.Column + 4
        R = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If R > lastRow Then lastRow = R
    Next c

    If lastRow < startCell.Row Then Exit Sub

    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(startCell, ws.Cells(lastRow, startCell.Column + 4))

    ' Limpiar contenido
    rng.ClearContents

    ' Eliminar todos los bordes
    rng.Borders(xlEdgeLeft).LineStyle = xlNone
    rng.Borders(xlEdgeTop).LineStyle = xlNone
    rng.Borders(xlEdgeBottom).LineStyle = xlNone
    rng.Borders(xlEdgeRight).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    On Error GoTo 0
End Sub