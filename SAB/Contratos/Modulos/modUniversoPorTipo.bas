' ========== mod3.bas  (TamañoPoblacion) ==========
Option Explicit

' ============================================================
'  Calcula UniversoPN, UniversoPJ y TamañoPob filtrando
'  la tabla Contratos por el período definido en la hoja Muestra.
'  Columnas reales del archivo: Fecha (type date), Tipo Persona.
'  PN = Tipo Persona empieza por "N" (Natural)
'  PJ = Tipo Persona empieza por "J" (Jurídico)
' ============================================================
Public Sub TamañoPoblacion()
    Dim wb As Workbook:       Set wb = ThisWorkbook
    Dim wsC As Worksheet, wsM As Worksheet
    Dim lo As ListObject
    Dim fechaCol As Long, tipoCol As Long
    Dim tipoInforme As String, mesTexto As String
    Dim anioFiltro As Long, mesFiltro As Long
    Dim db As Range
    Dim i As Long, total As Long, contN As Long, contJ As Long
    Dim fechaVal As Variant, tipoVal As String, initial As String

    On Error GoTo ErrHandler
    Application.EnableEvents   = False
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual

    ' --- Hojas y tabla ---
    On Error Resume Next
    Set wsC = wb.Worksheets("Contratos")
    Set wsM = wb.Worksheets("Muestra")
    On Error GoTo 0
    If wsC Is Nothing Or wsM Is Nothing Then GoTo Cleanup

    On Error Resume Next
    Set lo = wsC.ListObjects("Contratos")
    On Error GoTo 0
    If lo Is Nothing Then GoTo Cleanup
    If lo.DataBodyRange Is Nothing Then GoTo Cleanup

    ' --- Leer filtros desde la hoja Muestra ---
    On Error Resume Next
    tipoInforme = UCase$(Trim$(CStr(wb.Names("TipoInforme").RefersToRange.Value)))
    anioFiltro  = CLng(wb.Names("Año").RefersToRange.Value)
    If tipoInforme = "MENSUAL" Then
        mesTexto  = Trim$(CStr(wb.Names("Mes").RefersToRange.Value))
        mesFiltro = MesNumero(mesTexto)
    Else
        mesFiltro = 0   ' 0 = todos los meses del año
    End If
    On Error GoTo 0

    If anioFiltro = 0 Then GoTo Cleanup

    ' --- Localizar columnas ---
    fechaCol = ColIdx(lo, "Fecha")
    tipoCol  = ColIdx(lo, "Tipo Persona")

    If fechaCol = 0 Then
        MsgBox "No se encontró la columna 'Fecha' en la tabla 'Contratos'.", vbCritical
        GoTo Cleanup
    End If
    If tipoCol = 0 Then
        MsgBox "No se encontró la columna 'Tipo Persona' en la tabla 'Contratos'.", vbCritical
        GoTo Cleanup
    End If

    ' --- Contar ---
    Set db = lo.DataBodyRange
    total = 0: contN = 0: contJ = 0

    For i = 1 To db.Rows.Count
        fechaVal = db.Cells(i, fechaCol).Value
        If IsDate(fechaVal) Then
            If Year(fechaVal) = anioFiltro Then
                If mesFiltro = 0 Or Month(fechaVal) = mesFiltro Then
                    tipoVal = Trim$(CStr(db.Cells(i, tipoCol).Value))
                    If Len(tipoVal) > 0 Then
                        total = total + 1
                        initial = UCase$(Left$(tipoVal, 1))
                        If initial = "N" Then contN = contN + 1
                        If initial = "J" Then contJ = contJ + 1
                    End If
                End If
            End If
        End If
    Next i

    ' --- Guardar resultados ---
    On Error Resume Next
    wb.Names("TamañoPob").RefersToRange.Value   = total
    wb.Names("UniversoPN").RefersToRange.Value  = contN
    wb.Names("UniversoPJ").RefersToRange.Value  = contJ
    On Error GoTo 0

Cleanup:
    Application.Calculation    = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Exit Sub

ErrHandler:
    Application.Calculation    = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    MsgBox "Error en TamañoPoblacion: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ============================================================
'  HELPERS
' ============================================================

' Devuelve el índice (1-based) de una columna en un ListObject.
' Prueba coincidencia exacta primero, luego coincidencia parcial.
Public Function ColIdx(lo As ListObject, ByVal colName As String) As Long
    Dim i As Long, low As String, nm As String
    low = LCase$(colName)
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, colName, vbTextCompare) = 0 Then
            ColIdx = i: Exit Function
        End If
    Next i
    For i = 1 To lo.ListColumns.Count
        nm = LCase$(lo.ListColumns(i).Name)
        If InStr(nm, low) > 0 Then
            ColIdx = i: Exit Function
        End If
    Next i
    ColIdx = 0
End Function

' Convierte nombre o abreviatura de mes en español a número (1-12).
Public Function MesNumero(ByVal s As String) As Long
    Select Case UCase$(Left$(Trim$(s) & "   ", 3))
        Case "ENE": MesNumero = 1
        Case "FEB": MesNumero = 2
        Case "MAR": MesNumero = 3
        Case "ABR": MesNumero = 4
        Case "MAY": MesNumero = 5
        Case "JUN": MesNumero = 6
        Case "JUL": MesNumero = 7
        Case "AGO": MesNumero = 8
        Case "SEP", "SET": MesNumero = 9
        Case "OCT": MesNumero = 10
        Case "NOV": MesNumero = 11
        Case "DIC": MesNumero = 12
        Case Else:  MesNumero = 0
    End Select
End Function
