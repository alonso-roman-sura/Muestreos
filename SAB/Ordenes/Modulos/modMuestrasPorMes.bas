' ========== modMuestrasPorMes.bas ==========
Option Explicit

' Reconstruye “Universo/Muestra por mes” SIN tocar formatos fuera del bloque.
' Empieza inmediatamente debajo de los títulos (fila 4 si tus títulos están en la 3).
Public Sub MuestrasPorMes_Rebuild()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsM As Worksheet, wsO As Worksheet
    Dim lo As ListObject, fechas As Range
    Dim arr, R As Long, k As Variant

    On Error Resume Next
    Set wsM = wb.Worksheets("Muestra")
    Set wsO = wb.Worksheets("Ordenes")
    On Error GoTo 0
    If wsM Is Nothing Or wsO Is Nothing Then Exit Sub

    On Error Resume Next
    Set lo = wsO.ListObjects("Ordenes")
    If lo Is Nothing And wsO.ListObjects.Count > 0 Then Set lo = wsO.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.ListColumns("Fecha").DataBodyRange Is Nothing Then Exit Sub

    Set fechas = lo.ListColumns("Fecha").DataBodyRange
    If WorksheetFunction.CountA(fechas) = 0 Then Exit Sub
    arr = fechas.Value

    ' Meses únicos YYYY-MM
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    For R = 1 To UBound(arr, 1)
        If IsDate(arr(R, 1)) Then
            Dim y As Integer, m As Integer
            y = Year(arr(R, 1)): m = Month(arr(R, 1))
            dict(y & "-" & Format$(m, "00")) = Array(y, m)
        End If
    Next R
    If dict.Count = 0 Then
        ' no meses => limpiar solo el bloque previo y salir
        ClearPreviousBlockExact wsM
        modUtil.StoreMuestrasEndRow wsM, wsM.Range("D3").Row ' guarda “3” como tope (nada debajo)
        Exit Sub
    End If

    ' Ordenar claves
    Dim keys() As String, i As Long, j As Long, tmp As String
    ReDim keys(0 To dict.Count - 1)
    i = 0
    For Each k In dict.keys
        keys(i) = CStr(k): i = i + 1
    Next k
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i

    ' Plantillas y filas de inicio (justo debajo de títulos)
    Dim rngUnivLbl As Range, rngUnivVal As Range
    Dim rngMuesLbl As Range, rngMuesVal As Range
    Set rngUnivLbl = wsM.Range("J3:M3")
    Set rngUnivVal = wsM.Range("N3")
    Set rngMuesLbl = wsM.Range("D3:G3")
    Set rngMuesVal = wsM.Range("H3")

    Dim startRowU As Long, startRowM As Long
    startRowU = rngUnivLbl.Row + 1
    startRowM = rngMuesLbl.Row + 1

    ' Limpia exactamente el bloque previo usando el fin guardado
    ClearPreviousBlockExact wsM

    Dim rowU As Long, rowM As Long
    rowU = startRowU
    rowM = startRowM

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim y0 As Long, m0 As Long, tagMes As String, etiqueta As String
    Dim nmUniv As String, nmMues As String, f As String

    For i = 0 To UBound(keys)
        y0 = dict(keys(i))(0)
        m0 = dict(keys(i))(1)
        tagMes = MesAbrevES(m0) & CStr(y0)          ' p.ej. Jul2025

        ' ===== UNIVERSO (J:M combinado, N valor) =====
        With wsM.Range(wsM.Cells(rowU, "J"), wsM.Cells(rowU, "M"))
            If .mergeCells Then .UnMerge
            .Value = vbNullString
            ApplyLikeTemplate rngUnivLbl, .Resize(1, rngUnivLbl.Columns.Count), True
            etiqueta = "Universo Mes " & (i + 1) & " - " & MesAbrevES(m0) & " " & y0
            .Cells(1, 1).Value = etiqueta
        End With
        nmUniv = "Universo" & tagMes
        With wsM.Cells(rowU, "N")
            .ClearContents
            ApplyLikeTemplate rngUnivVal, .Resize(1, 1), False
            On Error Resume Next
            ThisWorkbook.Names(nmUniv).Delete
            On Error GoTo 0
            ThisWorkbook.Names.Add Name:=nmUniv, refersTo:="=" & wsM.Name & "!$N$" & rowU
            f = "=CONTAR.SI.CONJUNTO(" & _
                "Ordenes[Fecha];"">=""&FECHA(" & y0 & ";" & m0 & ";1);" & _
                "Ordenes[Fecha];""<""&FIN.MES(FECHA(" & y0 & ";" & m0 & ";1);0)+1;" & _
                "Ordenes[NºOrden];""<>"")"
            .FormulaLocal = f
        End With

        ' ===== MUESTRA (D:G combinado, H valor) =====
        With wsM.Range(wsM.Cells(rowM, "D"), wsM.Cells(rowM, "G"))
            If .mergeCells Then .UnMerge
            .Value = vbNullString
            ApplyLikeTemplate rngMuesLbl, .Resize(1, rngMuesLbl.Columns.Count), True
            etiqueta = "Tamaño de la muestra Mes " & (i + 1) & " - " & MesAbrevES(m0) & " " & y0
            .Cells(1, 1).Value = etiqueta
        End With
        nmMues = "Muestra" & tagMes
        With wsM.Cells(rowM, "H")
            .ClearContents
            ApplyLikeTemplate rngMuesVal, .Resize(1, 1), False
            On Error Resume Next
            ThisWorkbook.Names(nmMues).Delete
            On Error GoTo 0
            ThisWorkbook.Names.Add Name:=nmMues, refersTo:="=" & wsM.Name & "!$H$" & rowM
            .FormulaLocal = "=REDONDEAR.MAS((" & nmUniv & "*Z^2*p*(1-p))/(( " & nmUniv & " -1)*E^2+Z^2*p*(1-p));0)"
        End With

        rowU = rowU + 1
        rowM = rowM + 1
    Next i

    ' Guardar fin exacto del bloque dinámico para futuras limpiezas
    modUtil.StoreMuestrasEndRow wsM, Application.Max(rowU - 1, rowM - 1)

FIN:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Limpia SOLO el bloque previo definido por el nombre MuestrasEndRow (si existe)
Private Sub ClearPreviousBlockExact(ws As Worksheet)
    Dim rngUnivLbl As Range, rngMuesLbl As Range
    Set rngUnivLbl = ws.Range("J3:M3")
    Set rngMuesLbl = ws.Range("D3:G3")
    Dim startRowU As Long, startRowM As Long
    startRowU = rngUnivLbl.Row + 1
    startRowM = rngMuesLbl.Row + 1

    Dim prevEnd As Long
    prevEnd = modUtil.GetMuestrasEndRow(ws, Application.Max(startRowU, startRowM))

    If prevEnd < startRowU And prevEnd < startRowM Then Exit Sub

    Dim R As Long
    For R = startRowU To prevEnd
        With ws.Range(ws.Cells(R, "J"), ws.Cells(R, "M"))
            If .mergeCells Then .UnMerge
            .Value = vbNullString
        End With
        ws.Cells(R, "N").ClearContents
    Next R

    For R = startRowM To prevEnd
        With ws.Range(ws.Cells(R, "D"), ws.Cells(R, "G"))
            If .mergeCells Then .UnMerge
            .Value = vbNullString
        End With
        ws.Cells(R, "H").ClearContents
    Next R
End Sub

Private Function MesAbrevES(ByVal m As Long) As String
    Select Case m
        Case 1: MesAbrevES = "Ene"
        Case 2: MesAbrevES = "Feb"
        Case 3: MesAbrevES = "Mar"
        Case 4: MesAbrevES = "Abr"
        Case 5: MesAbrevES = "May"
        Case 6: MesAbrevES = "Jun"
        Case 7: MesAbrevES = "Jul"
        Case 8: MesAbrevES = "Ago"
        Case 9: MesAbrevES = "Sep"
        Case 10: MesAbrevES = "Oct"
        Case 11: MesAbrevES = "Nov"
        Case 12: MesAbrevES = "Dic"
        Case Else: MesAbrevES = "Mes"
    End Select
End Function