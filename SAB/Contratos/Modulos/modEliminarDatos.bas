' ========== modEliminarDatos.bas ==========
Option Explicit

Private Const HOJAS_PROTEGIDAS As String = "Instrucciones|Muestra"

' ============================================================
'  ENTRADA DEL BOTÓN "Eliminar Datos"
' ============================================================
Public Sub EliminarDatos()

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "Esta acción limpiará el archivo por completo:" & vbCrLf & vbCrLf & _
        "   " & Chr(149) & "  Todas las hojas excepto Instrucciones y Muestra" & vbCrLf & _
        "   " & Chr(149) & "  Todas las consultas y conexiones Power Query" & vbCrLf & _
        "   " & Chr(149) & "  Los valores generados en la hoja Muestra" & vbCrLf & vbCrLf & _
        "¿Desea continuar?", _
        vbYesNo + vbExclamation + vbDefaultButton2, "Eliminar datos")
    If resp <> vbYes Then Exit Sub

    resp = MsgBox( _
        "Esta operación no se puede deshacer." & vbCrLf & vbCrLf & _
        "¿Confirma que desea limpiar el archivo?", _
        vbYesNo + vbCritical + vbDefaultButton2, "Confirmar eliminación")
    If resp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo FIN

    EliminarHojas
    EliminarConexionesYConsultas
    LimpiarHojaMuestra

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "El archivo ha sido limpiado y está listo para recibir nuevos datos.", _
           vbInformation, "Listo"
    Exit Sub

FIN:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    If Err.Number <> 0 Then
        MsgBox "Error inesperado durante la limpieza:" & vbCrLf & Err.Description, _
               vbCritical, "Error"
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
            EsHojaProtegida = True
            Exit Function
        End If
    Next p
    EsHojaProtegida = False
End Function

' ============================================================
'  ELIMINAR CONSULTAS Y CONEXIONES POWER QUERY
' ============================================================
Private Sub EliminarConexionesYConsultas()
    Do While ThisWorkbook.Queries.Count > 0
        ThisWorkbook.Queries(1).Delete
    Loop

    Dim i As Long
    For i = ThisWorkbook.Connections.Count To 1 Step -1
        On Error Resume Next
        ThisWorkbook.Connections(i).Delete
        On Error GoTo 0
    Next i
End Sub

' ============================================================
'  LIMPIAR HOJA MUESTRA
'
'  En Contratos los nombres definidos (UniversoPN, UniversoPJ,
'  TamañoPob, TamañoMuestraPN, TamañoMuestraPJ, Muestra1_PN,
'  Muestra1_PJ, Mes, Año, TipoInforme) apuntan a celdas fijas
'  de la hoja Muestra, no se crean dinámicamente, por lo que
'  NO se eliminan. Solo se limpian sus valores.
'  Los dropdowns (Mes, Año, TipoInforme) se conservan intactos.
' ============================================================
Private Sub LimpiarHojaMuestra()
    Dim wsM As Worksheet
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets("Muestra")
    On Error GoTo 0
    If wsM Is Nothing Then Exit Sub

    ' Limpiar celdas de universo y tamaño de población
    LimpiarCeldaNombre "TamañoPob"
    LimpiarCeldaNombre "UniversoPN"
    LimpiarCeldaNombre "UniversoPJ"

    ' TamañoMuestraPN y TamañoMuestraPJ son fórmulas que dependen de
    ' UniversoPN/PJ; se limpian para no dejar errores residuales
    LimpiarCeldaNombre "TamañoMuestraPN"
    LimpiarCeldaNombre "TamañoMuestraPJ"

    ' Limpiar grillas de números aleatorios
    LimpiarGrillaMuestra "Muestra1_PN", 5
    LimpiarGrillaMuestra "Muestra1_PJ", 5
End Sub

' Vacía solo el contenido de la celda a la que apunta un nombre definido.
' Conserva el formato. No hace nada si el nombre no existe.
Private Sub LimpiarCeldaNombre(ByVal nmNombre As String)
    Dim nm As name, cel As Range
    On Error Resume Next
    Set nm = ThisWorkbook.Names(nmNombre)
    On Error GoTo 0
    If nm Is Nothing Then Exit Sub
    On Error Resume Next
    Set cel = nm.RefersToRange
    On Error GoTo 0
    If Not cel Is Nothing Then cel.ClearContents
End Sub

' Borra la grilla de números de muestra que empieza en la celda
' apuntada por nmNombre (nCols columnas de ancho, hacia abajo).
' Usa .Clear para eliminar también el formato punteado.
Private Sub LimpiarGrillaMuestra(ByVal nmNombre As String, ByVal nCols As Long)
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

    ' Determinar la última fila con datos en el bloque
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