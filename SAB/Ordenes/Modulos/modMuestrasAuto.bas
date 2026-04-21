Option Explicit

' ============================
'  MÓDULO: modMuestrasAuto
'  Seguro con formatos (usa modUtil.*)
' ============================

' === Punto de entrada (llamar tras importar Ordenes) ===
Public Sub PostOrdenes_ConstruirMuestras()
    Dim wsM As Worksheet
    Set wsM = modUtil.EnsureSheet("Muestra")

    ' Universo general en N3 y nombre "Universo"
    With wsM.Range("N3")
        .FormulaLocal = "=CONTARA(Ordenes[NºOrden])"
        modUtil.SafeDefineName "Universo", .Address(True, True, xlA1, True)
    End With

    ' Construir pares Universo/Muestra por mes
    BuildMuestrasPorMes
End Sub

' ============================
'   Núcleo
' ============================
Private Sub BuildMuestrasPorMes()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsM As Worksheet, wsO As Worksheet
    Dim lo As ListObject, fechas As Range
    Dim arr, R As Long

    Set wsM = modUtil.EnsureSheet("Muestra")

    On Error Resume Next
    Set wsO = wb.Worksheets("Ordenes")
    On Error GoTo 0
    If wsO Is Nothing Then Exit Sub

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
    If dict.Count = 0 Then Exit Sub

    ' Ordenar claves
    Dim keys() As String, i As Long, j As Long, tmp As String, k As Variant
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

    ' Plantillas (fila 3)
    Dim tplUnivLbl As Range, tplUnivVal As Range
    Dim tplMuesLbl As Range, tplMuesVal As Range
    Set tplUnivLbl = wsM.Range("J3:M3")
    Set tplUnivVal = wsM.Range("N3")
    Set tplMuesLbl = wsM.Range("D3:G3")
    Set tplMuesVal = wsM.Range("H3")

    ' Filas de inicio: justo debajo de los títulos
    Dim startRowU As Long, startRowM As Long
    startRowU = tplUnivLbl.Row + 1
    startRowM = tplMuesLbl.Row + 1

    ' Limpieza NO destructiva del bloque previo (solo contenidos)
    ClearPrevBlockByPrefix wsM, startRowU, "J", "M", "N", "Universo Mes "
    ClearPrevBlockByPrefix wsM, startRowM, "D", "G", "H", "Tamaño de la muestra Mes "

    ' Escritura pegada debajo
    Dim rowU As Long, rowM As Long
    rowU = startRowU
    rowM = startRowM

    Dim y0 As Long, m0 As Long, tagMes As String, etiqueta As String
    Dim nmUniv As String, nmMues As String, f As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 0 To UBound(keys)
        y0 = dict(keys(i))(0)
        m0 = dict(keys(i))(1)
        tagMes = MesAbrevES(m0) & CStr(y0)   ' p.ej. "Jul2025"

        ' ===== UNIVERSO (J:M combinado, N valor) =====
        With wsM.Range(wsM.Cells(rowU, ColNum("J")), wsM.Cells(rowU, ColNum("M")))
            If .mergeCells Then .UnMerge
            .Value = vbNullString
            modUtil.ApplyLikeTemplate tplUnivLbl, .Resize(1, tplUnivLbl.Columns.Count), True
            etiqueta = "Universo Mes " & (i + 1) & " - " & MesAbrevES(m0) & " " & y0
            .Cells(1, 1).Value = etiqueta
        End With

        nmUniv = "Universo" & tagMes
        With wsM.Cells(rowU, ColNum("N"))
            .ClearContents
            modUtil.ApplyLikeTemplate tplUnivVal, .Resize(1, 1), False
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
        With wsM.Range(wsM.Cells(rowM, ColNum("D")), wsM.Cells(rowM, ColNum("G")))
            If .mergeCells Then .UnMerge
            .Value = vbNullString
            modUtil.ApplyLikeTemplate tplMuesLbl, .Resize(1, tplMuesLbl.Columns.Count), True
            etiqueta = "Tamaño de la muestra Mes " & (i + 1) & " - " & MesAbrevES(m0) & " " & y0
            .Cells(1, 1).Value = etiqueta
        End With

        nmMues = "Muestra" & tagMes
        With wsM.Cells(rowM, ColNum("H"))
            .ClearContents
            modUtil.ApplyLikeTemplate tplMuesVal, .Resize(1, 1), False
            On Error Resume Next
            ThisWorkbook.Names(nmMues).Delete
            On Error GoTo 0
            ThisWorkbook.Names.Add Name:=nmMues, refersTo:="=" & wsM.Name & "!$H$" & rowM

            .FormulaLocal = "=REDONDEAR.MAS((" & nmUniv & "*Z^2*p*(1-p))/(( " & nmUniv & " -1)*E^2+Z^2*p*(1-p));0)"
        End With

        rowU = rowU + 1
        rowM = rowM + 1
    Next i

FIN:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ============================
'   Helpers locales
' ============================

' Borra SOLO filas previas del bloque dinámico cuyo label empiece con "prefix".
Private Sub ClearPrevBlockByPrefix(ws As Worksheet, ByVal startRow As Long, _
                                   ByVal colIni As String, ByVal colFin As String, _
                                   ByVal colVal As String, ByVal prefix As String)
    Dim R As Long, cIni As Long, cFin As Long, cVal As Long
    cIni = ColNum(colIni): cFin = ColNum(colFin): cVal = ColNum(colVal)

    R = startRow
    Do While True
        Dim lblCell As Range
        Set lblCell = ws.Cells(R, cIni)
        If IsEmpty(lblCell.Value) Then Exit Do
        If Left$(CStr(lblCell.Value), Len(prefix)) <> prefix Then Exit Do

        With ws.Range(ws.Cells(R, cIni), ws.Cells(R, cFin))
            If .mergeCells Then .UnMerge
            .Value = vbNullString          ' solo contenido del label
        End With
        ws.Cells(R, cVal).ClearContents    ' solo valor
        R = R + 1
    Loop
End Sub

' Convierte "A"?1, "N"?14, o deja números tal cual.
Private Function ColNum(ByVal colRef As Variant) As Long
    If IsNumeric(colRef) Then
        ColNum = CLng(colRef)
    Else
        ColNum = Columns(CStr(colRef)).Column
    End If
End Function

' Abreviatura de meses (uso local)
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
