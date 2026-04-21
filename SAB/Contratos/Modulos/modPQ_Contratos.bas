' ========== modPQ_Contratos.bas ==========
Option Explicit

' ============================================================
'  ENTRADA DEL BOTÓN "Importar Datos"
' ============================================================
Public Sub ImportarDatos()
    CargarContratos_PQ
End Sub

Public Sub CargarContratos_PQ()
    Dim ruta As String
    Dim ws As Worksheet

    ruta = PickFilePath()
    If Len(ruta) = 0 Then Exit Sub

    On Error GoTo ERR_HANDLER
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.CutCopyMode = False

    ResetContratosEnvironment ws
    UpsertQuery "Contratos", M_Contratos_PQ(ruta)

    Dim connStr As String, cmdText As String
    connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Contratos;Extended Properties=" & Chr(34) & Chr(34)
    cmdText = "SELECT * FROM [Contratos]"

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:=connStr, Destination:=ws.Range("A1"))
    With lo
        .name = "Contratos"
        With .QueryTable
            .CommandType = xlCmdSql
            .CommandText = cmdText
            .AdjustColumnWidth = True
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = False
            .MaintainConnection = True
            .RefreshStyle = xlOverwriteCells
            .SaveData = False
            .Refresh
        End With
    End With

    ' Restaurar eventos antes de la advertencia interactiva
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    AutodetectarPeriodo lo
    TamañoPoblacion

    MsgBox "Consulta 'Contratos' cargada correctamente.", vbInformation
    Exit Sub

ERR_HANDLER:
    Dim errDesc As String: errDesc = Err.Description
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error al cargar los datos:" & vbCrLf & vbCrLf & errDesc, _
           vbCritical, "Error de importaci" & Chr(243) & "n"
End Sub

' ============================================================
'  RESET
' ============================================================
Private Sub ResetContratosEnvironment(ByRef wsOut As Worksheet)
    Dim ws As Worksheet, lo As ListObject, cN As WorkbookConnection

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Do While ws.QueryTables.Count > 0
            If InStr(1, ws.QueryTables(1).Connection, "Location=Contratos", vbTextCompare) > 0 _
               Or InStr(1, ws.QueryTables(1).Connection, "Microsoft.Mashup.OleDb.1", vbTextCompare) > 0 Then
                ws.QueryTables(1).Delete
            Else
                Exit Do
            End If
        Loop
        Set lo = Nothing
        Set lo = ws.ListObjects("Contratos")
        If Not lo Is Nothing Then lo.Unlist
    Next ws

    For Each cN In ThisWorkbook.Connections
        If cN.Type = xlConnectionTypeOLEDB Then
            If InStr(1, cN.OLEDBConnection.Connection, "Location=Contratos", vbTextCompare) > 0 Then
                cN.Delete
            End If
        End If
    Next cN

    ThisWorkbook.Queries("Contratos").Delete
    On Error GoTo 0

    Set wsOut = Nothing
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Contratos")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsOut.name = "Contratos"
    Else
        wsOut.Cells.Clear
    End If
End Sub

Private Sub UpsertQuery(ByVal qName As String, ByVal mCode As String)
    Dim q As WorkbookQuery
    On Error Resume Next
    Set q = ThisWorkbook.Queries(qName)
    On Error GoTo 0
    If q Is Nothing Then
        ThisWorkbook.Queries.Add name:=qName, Formula:=mCode
    Else
        q.Formula = mCode
    End If
End Sub

Private Function PickFilePath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccionar archivo de Contratos (.XLS, .CSV, .TXT)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Archivos comunes", "*.xls; *.xlsx; *.xlsm; *.xlsb; *.csv; *.txt"
        If .Show <> -1 Then Exit Function
        PickFilePath = .SelectedItems(1)
    End With
End Function

' ============================================================
'  DETECCIÓN DE PERÍODO
'  Detecta los meses presentes en los datos importados.
'  Si hay más de un mes, advierte al usuario y pregunta si
'  desea proceder de todas formas.
'  Escribe la etiqueta de período en la celda PeriodoActual.
' ============================================================
Private Sub AutodetectarPeriodo(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim fechaCol As Long
    fechaCol = ColIdx(lo, "Fecha de Ingreso")
    If fechaCol = 0 Then Exit Sub

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim db As Range: Set db = lo.DataBodyRange
    Dim i As Long, fv As Variant, y As Long, m As Long

    For i = 1 To db.Rows.Count
        fv = db.Cells(i, fechaCol).Value
        If IsDate(fv) Then
            y = Year(fv): m = Month(fv)
            dict(y & "-" & Format$(m, "00")) = Array(y, m)
        End If
    Next i

    If dict.Count = 0 Then Exit Sub

    ' Advertencia si hay más de un mes
    If dict.Count > 1 Then
        Dim lista As String, k As Variant
        For Each k In dict.keys
            Dim arr As Variant: arr = dict(k)
            lista = lista & "  " & Chr(149) & "  " & NombreMesES(arr(1)) & " " & arr(0) & vbCrLf
        Next k
        Dim resp As VbMsgBoxResult
        resp = MsgBox( _
            "El archivo contiene " & dict.Count & " meses distintos:" & vbCrLf & vbCrLf & _
            lista & vbCrLf & _
            "Se recomienda importar archivos de un solo mes." & vbCrLf & vbCrLf & _
            "¿Desea continuar de todas formas?", _
            vbYesNo + vbExclamation + vbDefaultButton2, _
            "Archivo con m" & Chr(250) & "ltiples meses")
        If resp <> vbYes Then Exit Sub
    End If

    ' Construir etiqueta de período
    Dim etiqueta As String
    If dict.Count = 1 Then
        Dim arr1 As Variant: arr1 = dict.Items()(0)
        etiqueta = NombreMesES(arr1(1)) & " " & arr1(0)
    Else
        ' Rango: primer mes - último mes del año más reciente
        Dim keys() As String
        ReDim keys(0 To dict.Count - 1)
        i = 0
        For Each k In dict.keys
            keys(i) = CStr(k): i = i + 1
        Next k
        ' Ordenar
        Dim j As Long, tmp As String
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If keys(j) < keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            Next j
        Next i
        Dim arrFirst As Variant: arrFirst = dict(keys(0))
        Dim arrLast  As Variant: arrLast = dict(keys(UBound(keys)))
        etiqueta = NombreMesES(arrFirst(1)) & " " & arrFirst(0) & _
                   " - " & NombreMesES(arrLast(1)) & " " & arrLast(0)
    End If

    On Error Resume Next
    wb.Names("PeriodoActual").RefersToRange.Value = etiqueta
    On Error GoTo 0
End Sub

' ============================================================
'  HELPERS
' ============================================================

Private Function ColIdx(lo As ListObject, ByVal colName As String) As Long
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

Private Function NombreMesES(ByVal m As Long) As String
    Select Case m
        Case 1:  NombreMesES = "Enero"
        Case 2:  NombreMesES = "Febrero"
        Case 3:  NombreMesES = "Marzo"
        Case 4:  NombreMesES = "Abril"
        Case 5:  NombreMesES = "Mayo"
        Case 6:  NombreMesES = "Junio"
        Case 7:  NombreMesES = "Julio"
        Case 8:  NombreMesES = "Agosto"
        Case 9:  NombreMesES = "Septiembre"
        Case 10: NombreMesES = "Octubre"
        Case 11: NombreMesES = "Noviembre"
        Case 12: NombreMesES = "Diciembre"
        Case Else: NombreMesES = ""
    End Select
End Function

' ============================================================
'  M FORMULA
'  27 columnas reales del archivo de Contratos.
'  Columnas clave: "Tipo" (N/J), "Fecha de Ingreso" (date).
'  Sort: Fecha de Ingreso ASC, Cuenta ASC.
' ============================================================
Private Function M_Contratos_PQ(ByVal ruta As String) As String
    Dim m As String, p As String
    p = Replace(ruta, """", """""")

    Dim cClasif1  As String: cClasif1 = "Clasificaci" & Chr(243) & "n 1"
    Dim cClasif2  As String: cClasif2 = "Clasificaci" & Chr(243) & "n 2"
    Dim cDirPrec  As String: cDirPrec = "Direcci" & Chr(243) & "n Precisa"
    Dim cDirCont  As String: cDirCont = "Direcci" & Chr(243) & "n de Contacto"
    Dim cTelefono As String: cTelefono = "Tel" & Chr(233) & "fono"
    Dim cPais     As String: cPais = "Pa" & Chr(237) & "s"
    Dim cEnvio    As String: cEnvio = "Lugar de Env" & Chr(237) & "o de Correspondencia"
    Dim cFechaIng As String: cFechaIng = "Fecha de Ingreso"
    Dim cFecBloq  As String: cFecBloq = "Fecha de Bloqueo"

    m = "let" & vbCrLf
    m = m & "  path = """ & p & """," & vbCrLf & vbCrLf

    m = m & "  expected = {""Cuenta"",""Tipo"",""Nombre"",""RUC/NIT""," & vbCrLf
    m = m & "               """ & cClasif1 & """,""" & cClasif2 & """," & vbCrLf
    m = m & "               """ & cDirPrec & """,""" & cDirCont & """," & vbCrLf
    m = m & "               """ & cTelefono & """,""Celular"",""Fax"",""Casilla"",""Email""," & vbCrLf
    m = m & "               """ & cEnvio & """," & vbCrLf
    m = m & "               ""Oficial de Cuenta"",""Referencia"",""" & cFechaIng & """,""" & cPais & """,""Distrito""," & vbCrLf
    m = m & "               ""C Entero"",""Conoc Merc"",""Estado"",""Tipo Bloqueo"",""" & cFecBloq & """," & vbCrLf
    m = m & "               ""Observaciones del Agente"",""Tipo de Cliente"",""Vinculado a Agente""}," & vbCrLf & vbCrLf

    m = m & "  Canon = (s as text) as text =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      t0 = Text.Upper(Text.Trim(s))," & vbCrLf
    m = m & "      t1 = Text.Replace(t0,""" & Chr(193) & """,""A"")," & vbCrLf
    m = m & "      t2 = Text.Replace(t1,""" & Chr(201) & """,""E"")," & vbCrLf
    m = m & "      t3 = Text.Replace(t2,""" & Chr(205) & """,""I"")," & vbCrLf
    m = m & "      t4 = Text.Replace(t3,""" & Chr(211) & """,""O"")," & vbCrLf
    m = m & "      t5 = Text.Replace(t4,""" & Chr(218) & """,""U"")," & vbCrLf
    m = m & "      t6 = Text.Replace(t5,""" & Chr(209) & """,""N"")," & vbCrLf
    m = m & "      t7 = Text.Replace(Text.Replace(t6,""N" & Chr(186) & """,""N""),""N" & Chr(176) & """,""N"")," & vbCrLf
    m = m & "      t8 = Text.Replace("" "" & t7 & "" "","" DE "","" "")," & vbCrLf
    m = m & "      out = Text.Remove(Text.Trim(t8), {"" "",""_"",""-"",""."",""/"",""\""})" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      out," & vbCrLf & vbCrLf

    m = m & "  expectedCanon = List.Transform(expected, each Canon(_))," & vbCrLf
    m = m & "  lenExp = List.Count(expected)," & vbCrLf & vbCrLf

    m = m & "  bin = Binary.Buffer(File.Contents(path))," & vbCrLf
    m = m & "  encodings = {65001, 1252}," & vbCrLf
    m = m & "  delims    = {" & Chr(34) & "," & Chr(34) & "," & Chr(34) & "#(tab)" & Chr(34) & "," & Chr(34) & "|" & Chr(34) & "}," & vbCrLf & vbCrLf

    m = m & "  MakeTable = (enc as number, delim as text) as nullable table =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      t = try Csv.Document(bin, [Delimiter=delim, Columns=null, Encoding=enc, QuoteStyle=QuoteStyle.Csv]) otherwise null" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      if t = null then null else t," & vbCrLf & vbCrLf

    m = m & "  RowIsExactHeader = (row as list) as logical =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      rTxt  = List.Transform(row, each if _ = null then """" else Text.From(_))," & vbCrLf
    m = m & "      rCan  = List.Transform(rTxt, each Canon(_))," & vbCrLf
    m = m & "      okLen = List.Count(rCan) >= lenExp," & vbCrLf
    m = m & "      slice = if okLen then List.FirstN(rCan, lenExp) else rCan," & vbCrLf
    m = m & "      eqAll = okLen and List.AllTrue(List.Transform({0..lenExp-1}, each slice{_} = expectedCanon{_}))" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      eqAll," & vbCrLf & vbCrLf

    m = m & "  FindHeaderRowIndex = (tbl as table, maxScan as number) as any =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      sample = Table.FirstN(tbl, maxScan)," & vbCrLf
    m = m & "      rows   = Table.ToRows(sample)," & vbCrLf
    m = m & "      idx    = List.PositionOf(List.Transform(rows, each RowIsExactHeader(_)), true)" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      idx," & vbCrLf & vbCrLf

    m = m & "  TryAll =" & vbCrLf
    m = m & "    List.First(" & vbCrLf
    m = m & "      List.RemoveNulls(" & vbCrLf
    m = m & "        List.Transform(" & vbCrLf
    m = m & "          encodings," & vbCrLf
    m = m & "          (enc) =>" & vbCrLf
    m = m & "            List.First(" & vbCrLf
    m = m & "              List.RemoveNulls(" & vbCrLf
    m = m & "                List.Transform(" & vbCrLf
    m = m & "                  delims," & vbCrLf
    m = m & "                  (delim) =>" & vbCrLf
    m = m & "                    let" & vbCrLf
    m = m & "                      t0  = MakeTable(enc, delim)," & vbCrLf
    m = m & "                      idx = if t0 = null then -1 else FindHeaderRowIndex(t0, 120)," & vbCrLf
    m = m & "                      rec = if t0 <> null and idx >= 0 then [Enc=enc, Delim=delim, Tbl=t0, HeaderIdx=idx] else null" & vbCrLf
    m = m & "                    in" & vbCrLf
    m = m & "                      rec" & vbCrLf
    m = m & "                )" & vbCrLf
    m = m & "              )" & vbCrLf
    m = m & "            )" & vbCrLf
    m = m & "        )" & vbCrLf
    m = m & "      )," & vbCrLf
    m = m & "      null" & vbCrLf
    m = m & "    )," & vbCrLf & vbCrLf

    m = m & "  tblAll = if TryAll = null" & vbCrLf
    m = m & "           then error ""No se encontr" & Chr(243) & " la fila con las 27 cabeceras esperadas en el archivo.""" & vbCrLf
    m = m & "           else TryAll[Tbl]," & vbCrLf
    m = m & "  hIdx   = TryAll[HeaderIdx]," & vbCrLf & vbCrLf

    m = m & "  tblAfterSkip = Table.Skip(tblAll, hIdx)," & vbCrLf
    m = m & "  promoted     = Table.PromoteHeaders(tblAfterSkip, [PromoteAllScalars=true])," & vbCrLf & vbCrLf

    m = m & "  curNames = Table.ColumnNames(promoted)," & vbCrLf
    m = m & "  pairs    = List.Transform({0..lenExp-1}, each {curNames{_}, expected{_}})," & vbCrLf
    m = m & "  renamed  = Table.RenameColumns(promoted, pairs, MissingField.Ignore)," & vbCrLf & vbCrLf

    m = m & "  only27   = Table.SelectColumns(renamed, expected, MissingField.UseNull)," & vbCrLf & vbCrLf

    m = m & "  allTextCols = List.RemoveItems(expected, {""" & cFechaIng & """,""" & cFecBloq & """})," & vbCrLf
    m = m & "  AsText = Table.TransformColumnTypes(only27, List.Transform(allTextCols, each {_, type text}), ""es-PE"")," & vbCrLf & vbCrLf

    m = m & "  MonthPairs = {" & vbCrLf
    m = m & "    ""ene"",""enero"",""feb"",""febrero"",""mar"",""marzo"",""abr"",""abril"",""may"",""mayo"",""jun"",""junio""," & vbCrLf
    m = m & "    ""jul"",""julio"",""ago"",""agosto"",""set"",""septiembre"",""sep"",""septiembre"",""sept"",""septiembre""," & vbCrLf
    m = m & "    ""oct"",""octubre"",""nov"",""noviembre"",""dic"",""diciembre""" & vbCrLf
    m = m & "  }," & vbCrLf & vbCrLf

    m = m & "  NormalizeMonthText = (t as text) as text =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      s0 = Text.Trim(t)," & vbCrLf
    m = m & "      s1 = Text.Lower(s0)," & vbCrLf
    m = m & "      validMons = {""ene"",""feb"",""mar"",""abr"",""may"",""jun"",""jul"",""ago"",""set"",""sep"",""oct"",""nov"",""dic""}," & vbCrLf
    m = m & "      mon2 = Text.Middle(s1, 2, 3)," & vbCrLf
    m = m & "      yr2  = Text.End(s1, Text.Length(s1) - 5)," & vbCrLf
    m = m & "      ok2  = Text.Length(s1) >= 7 and (try Number.From(Text.Start(s1,2)) >= 1 otherwise false)" & vbCrLf
    m = m & "             and List.Contains(validMons, mon2) and (try Number.From(yr2) >= 1 otherwise false)," & vbCrLf
    m = m & "      mon1 = Text.Middle(s1, 1, 3)," & vbCrLf
    m = m & "      yr1  = Text.End(s1, Text.Length(s1) - 4)," & vbCrLf
    m = m & "      ok1  = Text.Length(s1) >= 6 and (try Number.From(Text.Start(s1,1)) >= 1 otherwise false)" & vbCrLf
    m = m & "             and List.Contains(validMons, mon1) and (try Number.From(yr1) >= 1 otherwise false)," & vbCrLf
    m = m & "      dayPart = if ok2 then Text.Start(s1, 2) else if ok1 then Text.Start(s1, 1) else """"," & vbCrLf
    m = m & "      monPart = if ok2 then mon2 else if ok1 then mon1 else """"," & vbCrLf
    m = m & "      yrRaw   = if ok2 then yr2 else if ok1 then yr1 else """"," & vbCrLf
    m = m & "      yrNorm  = if Text.Length(yrRaw) = 2 then" & vbCrLf
    m = m & "                  (if Number.FromText(yrRaw) < 50 then ""20"" else ""19"") & yrRaw" & vbCrLf
    m = m & "                else if Text.Length(yrRaw) = 3 then" & vbCrLf
    m = m & "                  ""20"" & Text.End(yrRaw, 2)" & vbCrLf
    m = m & "                else yrRaw," & vbCrLf
    m = m & "      sExp    = if ok2 or ok1 then dayPart & "" "" & monPart & "" "" & yrNorm else s1," & vbCrLf
    m = m & "      s2 = Text.Replace(Text.Replace(Text.Replace(sExp, ""/"", "" ""), ""-"", "" ""), ""."", """")," & vbCrLf
    m = m & "      parts = List.Select(Text.Split(s2, "" ""), each _ <> """")," & vbCrLf
    m = m & "      s3 = Text.Combine(parts, "" "")," & vbCrLf
    m = m & "      s4 = "" "" & s3 & "" ""," & vbCrLf
    m = m & "      s5 = List.Accumulate({0..List.Count(MonthPairs)/2-1}, s4, (state, i) => Text.Replace(state, "" "" & MonthPairs{2*i} & "" "", "" "" & MonthPairs{2*i+1} & "" ""))," & vbCrLf
    m = m & "      res = Text.Trim(s5)" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      res," & vbCrLf & vbCrLf

    m = m & "  ParseFechaES = (v as any) as nullable date =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      out =" & vbCrLf
    m = m & "        if v is date then Date.From(v) else" & vbCrLf
    m = m & "        if v is datetime then Date.From(v) else" & vbCrLf
    m = m & "        if v is number then Date.From(#datetime(1899,12,30,0,0,0) + #duration(Number.From(v),0,0,0)) else" & vbCrLf
    m = m & "        if v is text then" & vbCrLf
    m = m & "          let" & vbCrLf
    m = m & "            n   = NormalizeMonthText(v)," & vbCrLf
    m = m & "            d1  = try Date.FromText(n, ""es-PE"") otherwise try Date.FromText(n, ""es-ES"") otherwise null," & vbCrLf
    m = m & "            res = if d1 <> null then d1 else" & vbCrLf
    m = m & "                    let" & vbCrLf
    m = m & "                      parts = Text.Split(n, "" "")," & vbCrLf
    m = m & "                      dS = if List.Count(parts) > 0 then parts{0} else """"," & vbCrLf
    m = m & "                      mS = if List.Count(parts) > 1 then parts{1} else """"," & vbCrLf
    m = m & "                      yS = if List.Count(parts) > 2 then parts{2} else """"," & vbCrLf
    m = m & "                      dN = try Number.FromText(dS) otherwise null," & vbCrLf
    m = m & "                      months = {""enero"",""febrero"",""marzo"",""abril"",""mayo"",""junio"",""julio"",""agosto"",""septiembre"",""octubre"",""noviembre"",""diciembre""}," & vbCrLf
    m = m & "                      mPos = List.PositionOf(months, mS)," & vbCrLf
    m = m & "                      mN = if mPos >= 0 then mPos + 1 else null," & vbCrLf
    m = m & "                      yN0 = try Number.FromText(yS) otherwise null," & vbCrLf
    m = m & "                      yN  = if yN0 is number and Text.Length(yS)=2 then (if yN0 < 50 then 2000 + yN0 else 1900 + yN0) else yN0," & vbCrLf
    m = m & "                      d2 = if dN <> null and mN <> null and yN <> null then #date(yN, mN, dN) else null" & vbCrLf
    m = m & "                    in" & vbCrLf
    m = m & "                      d2" & vbCrLf
    m = m & "          in" & vbCrLf
    m = m & "            res" & vbCrLf
    m = m & "        else" & vbCrLf
    m = m & "          null" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      out," & vbCrLf & vbCrLf

    m = m & "  FechaIngFix = Table.TransformColumns(AsText, {{""" & cFechaIng & """, each ParseFechaES(_), type date}})," & vbCrLf
    m = m & "  FechaFix    = Table.TransformColumns(FechaIngFix, {{""" & cFecBloq & """, each ParseFechaES(_), type date}})," & vbCrLf & vbCrLf

    m = m & "  NoNulls = Table.SelectRows(FechaFix, each Record.Field(_, """ & cFechaIng & """) <> null)," & vbCrLf & vbCrLf

    m = m & "  Sorted = Table.Sort(NoNulls, {{""" & cFechaIng & """, Order.Ascending}, {""Cuenta"", Order.Ascending}})" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "  Sorted" & vbCrLf

    M_Contratos_PQ = m
End Function