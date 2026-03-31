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

    On Error GoTo FIN
    Application.ScreenUpdating = False
    Application.EnableEvents  = False
    Application.DisplayAlerts = False
    Application.CutCopyMode  = False

    ResetContratosEnvironment ws
    UpsertQuery "Contratos", M_Contratos_PQ(ruta)

    Dim connStr As String, cmdText As String
    connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Contratos;Extended Properties=" & Chr(34) & Chr(34)
    cmdText = "SELECT * FROM [Contratos]"

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:=connStr, Destination:=ws.Range("A1"))
    With lo
        .Name = "Contratos"
        With .QueryTable
            .CommandType    = xlCmdSql
            .CommandText    = cmdText
            .AdjustColumnWidth  = True
            .PreserveFormatting = True
            .RefreshOnFileOpen  = False
            .BackgroundQuery    = False
            .MaintainConnection = True
            .RefreshStyle       = xlOverwriteCells
            .SaveData           = False
            .Refresh
        End With
    End With

    ' Actualizar universos con los filtros actuales de la hoja Muestra
    On Error Resume Next
    TamañoPoblacion
    On Error GoTo 0

FIN:
    Application.DisplayAlerts  = True
    Application.EnableEvents   = True
    Application.ScreenUpdating = True

    If Not ws Is Nothing Then
        If ws.ListObjects.Count > 0 Then
            MsgBox "Consulta 'Contratos' cargada correctamente.", vbInformation
        End If
    End If
End Sub

' ============================================================
'  RESET: limpia hoja, QueryTables, conexiones y query previas
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
        wsOut.Name = "Contratos"
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
        ThisWorkbook.Queries.Add Name:=qName, Formula:=mCode
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
'  M FORMULA (Power Query)
'  Detecta automáticamente el encoding y delimitador.
'  Busca la fila con las 24 cabeceras esperadas (hasta fila 120).
'  Parsea Fecha (DDMMMYYYY) y ordena por Fecha ASC, Transac ASC.
' ============================================================
Private Function M_Contratos_PQ(ByVal ruta As String) As String
    Dim m As String, p As String
    p = Replace(ruta, """", """""")

    m = "let" & vbCrLf
    m = m & "  path = """ & p & """," & vbCrLf & vbCrLf

    m = m & "  expected = {""Transac"",""Fecha"",""Cuenta"",""Documento"",""Tipo Persona"",""Tipo Doc"",""OfCta""," & vbCrLf
    m = m & "               ""Como se Enter\u00f3"",""Ref"",""Moneda Ori"",""Monto Ori"",""Moneda Des"",""Monto Des""," & vbCrLf
    m = m & "               ""TC"",""TCBanco"",""MonExp"",""Total Neto"",""Mon G/P"",""Gan/Per"",""Gan/Per PEN""," & vbCrLf
    m = m & "               ""Cbte"",""Canal"",""Flujo en GAM"",""Confirmaci\u00f3n Correo""}," & vbCrLf & vbCrLf

    ' Reemplazar escapes unicode por caracteres reales
    m = Replace(m, "\u00f3", Chr(243))   ' ó
    m = Replace(m, "\u00f3", Chr(243))

    m = m & "  Canon = (s as text) as text =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      t0 = Text.Upper(Text.Trim(s))," & vbCrLf
    m = m & "      t1 = Text.Replace(t0,""Á"",""A"")," & vbCrLf
    m = m & "      t2 = Text.Replace(t1,""É"",""E"")," & vbCrLf
    m = m & "      t3 = Text.Replace(t2,""Í"",""I"")," & vbCrLf
    m = m & "      t4 = Text.Replace(t3,""Ó"",""O"")," & vbCrLf
    m = m & "      t5 = Text.Replace(t4,""Ú"",""U"")," & vbCrLf
    m = m & "      t6 = Text.Replace(t5,""Ñ"",""N"")," & vbCrLf
    m = m & "      out = Text.Remove(Text.Trim(t6), {"" "",""_"",""-"",""."",""/"",""\""})" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      out," & vbCrLf & vbCrLf

    m = m & "  lenExp = List.Count(expected)," & vbCrLf
    m = m & "  expectedCanon = List.Transform(expected, each Canon(_))," & vbCrLf & vbCrLf

    m = m & "  bin = Binary.Buffer(File.Contents(path))," & vbCrLf
    m = m & "  encodings = {65001, 1252}," & vbCrLf
    m = m & "  delims    = {""" & Chr(9) & """, "","", ""|""}," & vbCrLf & vbCrLf

    m = m & "  MakeTable = (enc as number, delim as text) as nullable table =>" & vbCrLf
    m = m & "    let t = try Csv.Document(bin, [Delimiter=delim, Columns=null, Encoding=enc, QuoteStyle=QuoteStyle.Csv]) otherwise null" & vbCrLf
    m = m & "    in if t = null then null else t," & vbCrLf & vbCrLf

    m = m & "  RowIsExactHeader = (row as list) as logical =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      rTxt  = List.Transform(row, each if _ = null then """" else Text.From(_))," & vbCrLf
    m = m & "      rCan  = List.Transform(rTxt, each Canon(_))," & vbCrLf
    m = m & "      okLen = List.Count(rCan) >= lenExp," & vbCrLf
    m = m & "      slice = if okLen then List.FirstN(rCan, lenExp) else rCan," & vbCrLf
    m = m & "      eqAll = okLen and List.AllTrue(List.Transform({0..lenExp-1}, each slice{_} = expectedCanon{_}))" & vbCrLf
    m = m & "    in eqAll," & vbCrLf & vbCrLf

    m = m & "  FindHeaderRowIndex = (tbl as table, maxScan as number) as any =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      rows = Table.ToRows(Table.FirstN(tbl, maxScan))," & vbCrLf
    m = m & "      idx  = List.PositionOf(List.Transform(rows, each RowIsExactHeader(_)), true)" & vbCrLf
    m = m & "    in idx," & vbCrLf & vbCrLf

    m = m & "  TryAll = List.First(List.RemoveNulls(List.Transform(encodings, (enc) =>" & vbCrLf
    m = m & "    List.First(List.RemoveNulls(List.Transform(delims, (delim) =>" & vbCrLf
    m = m & "      let t0  = MakeTable(enc, delim)," & vbCrLf
    m = m & "          idx = if t0 = null then -1 else FindHeaderRowIndex(t0, 120)," & vbCrLf
    m = m & "          rec = if t0 <> null and idx >= 0 then [Enc=enc, Delim=delim, Tbl=t0, HeaderIdx=idx] else null" & vbCrLf
    m = m & "      in rec))))), null)," & vbCrLf & vbCrLf

    m = m & "  _fail = if TryAll = null then error ""No se encontraron las 24 cabeceras esperadas."" else null," & vbCrLf & vbCrLf

    m = m & "  tblAll       = TryAll[Tbl]," & vbCrLf
    m = m & "  hIdx         = TryAll[HeaderIdx]," & vbCrLf
    m = m & "  tblAfterSkip = Table.Skip(tblAll, hIdx)," & vbCrLf
    m = m & "  promoted     = Table.PromoteHeaders(tblAfterSkip, [PromoteAllScalars=true])," & vbCrLf & vbCrLf

    m = m & "  curNames = Table.ColumnNames(promoted)," & vbCrLf
    m = m & "  pairs    = List.Transform({0..lenExp-1}, each {curNames{_}, expected{_}})," & vbCrLf
    m = m & "  renamed  = Table.RenameColumns(promoted, pairs, MissingField.Ignore)," & vbCrLf
    m = m & "  only24   = Table.SelectColumns(renamed, expected, MissingField.UseNull)," & vbCrLf & vbCrLf

    ' Fecha como date, todo lo demás como text
    m = m & "  allTextCols = List.RemoveItems(expected, {""Fecha""})," & vbCrLf
    m = m & "  AsText = Table.TransformColumnTypes(only24, List.Transform(allTextCols, each {_, type text}), ""es-PE"")," & vbCrLf & vbCrLf

    ' Normalización de mes en texto compacto DDMMMYYYY
    m = m & "  MonthPairs = {""ene"",""enero"",""feb"",""febrero"",""mar"",""marzo"",""abr"",""abril""," & vbCrLf
    m = m & "    ""may"",""mayo"",""jun"",""junio"",""jul"",""julio"",""ago"",""agosto""," & vbCrLf
    m = m & "    ""set"",""septiembre"",""sep"",""septiembre"",""sept"",""septiembre""," & vbCrLf
    m = m & "    ""oct"",""octubre"",""nov"",""noviembre"",""dic"",""diciembre""}," & vbCrLf & vbCrLf

    m = m & "  NormalizeMonthText = (t as text) as text =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      s0 = Text.Trim(t)," & vbCrLf
    m = m & "      s1 = Text.Lower(s0)," & vbCrLf
    ' Expansión DDMMMYYYY a "DD MMM YYYY"
    m = m & "      isC9 = (Text.Length(s1)=9) and (try (Number.From(Text.Start(s1,2))>=1) otherwise false) and (try (Number.From(Text.End(s1,4))>=1000) otherwise false)," & vbCrLf
    m = m & "      isC8 = (Text.Length(s1)=8) and (try (Number.From(Text.Start(s1,1))>=1) otherwise false) and (try (Number.From(Text.End(s1,4))>=1000) otherwise false)," & vbCrLf
    m = m & "      sExp = if isC9 then Text.Start(s1,2)&"" ""&Text.Middle(s1,2,3)&"" ""&Text.End(s1,4)" & vbCrLf
    m = m & "             else if isC8 then Text.Start(s1,1)&"" ""&Text.Middle(s1,1,3)&"" ""&Text.End(s1,4)" & vbCrLf
    m = m & "             else s1," & vbCrLf
    m = m & "      s2 = Text.Replace(Text.Replace(Text.Replace(sExp,""/"","" ""),""-"","" ""),""."","""")," & vbCrLf
    m = m & "      s3 = Text.Combine(List.Select(Text.Split(s2,"" ""), each _ <> """"),"" "")," & vbCrLf
    m = m & "      s4 = "" "" & s3 & "" ""," & vbCrLf
    m = m & "      s5 = List.Accumulate({0..List.Count(MonthPairs)/2-1}, s4, (state,i) => Text.Replace(state,"" ""&MonthPairs{2*i}&"" "","" ""&MonthPairs{2*i+1}&"" ""))," & vbCrLf
    m = m & "      res = Text.Trim(s5)" & vbCrLf
    m = m & "    in res," & vbCrLf & vbCrLf

    m = m & "  ParseFechaES = (v as any) as nullable date =>" & vbCrLf
    m = m & "    let out =" & vbCrLf
    m = m & "        if v is date then Date.From(v) else" & vbCrLf
    m = m & "        if v is datetime then Date.From(v) else" & vbCrLf
    m = m & "        if v is number then Date.From(#datetime(1899,12,30,0,0,0)+#duration(Number.From(v),0,0,0)) else" & vbCrLf
    m = m & "        if v is text then" & vbCrLf
    m = m & "          let n  = NormalizeMonthText(v)," & vbCrLf
    m = m & "              d1 = try Date.FromText(n,""es-PE"") otherwise try Date.FromText(n,""es-ES"") otherwise null," & vbCrLf
    m = m & "              res = if d1 <> null then d1 else" & vbCrLf
    m = m & "                      let parts = Text.Split(n,"" "")," & vbCrLf
    m = m & "                          dN = try Number.FromText(parts{0}) otherwise null," & vbCrLf
    m = m & "                          months = {""enero"",""febrero"",""marzo"",""abril"",""mayo"",""junio"",""julio"",""agosto"",""septiembre"",""octubre"",""noviembre"",""diciembre""}," & vbCrLf
    m = m & "                          mPos = if List.Count(parts)>1 then List.PositionOf(months, parts{1}) else -1," & vbCrLf
    m = m & "                          mN = if mPos >= 0 then mPos+1 else null," & vbCrLf
    m = m & "                          yN0 = if List.Count(parts)>2 then try Number.FromText(parts{2}) otherwise null else null," & vbCrLf
    m = m & "                          yN  = if yN0 is number and Text.Length(parts{2})=2 then (if yN0<50 then 2000+yN0 else 1900+yN0) else yN0" & vbCrLf
    m = m & "                      in if dN<>null and mN<>null and yN<>null then #date(yN,mN,dN) else null" & vbCrLf
    m = m & "          in res" & vbCrLf
    m = m & "        else null" & vbCrLf
    m = m & "    in out," & vbCrLf & vbCrLf

    m = m & "  FechaFix = Table.TransformColumns(AsText, {{""Fecha"", each ParseFechaES(_), type date}})," & vbCrLf & vbCrLf

    ' Eliminar filas sin fecha válida (filas de totales o vacías al final)
    m = m & "  NoNulls = Table.SelectRows(FechaFix, each [Fecha] <> null)," & vbCrLf & vbCrLf

    ' Ordenar por Fecha ASC, luego Transac ASC
    m = m & "  Sorted = Table.Sort(NoNulls, {{""Fecha"", Order.Ascending}, {""Transac"", Order.Ascending}})" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "  Sorted" & vbCrLf

    M_Contratos_PQ = m
End Function
