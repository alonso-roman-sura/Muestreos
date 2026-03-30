' ========== modPQ_Ordenes.bas ==========
Option Explicit

Public Sub ImportarDatos()
    CargarOrdenes_PQ
End Sub

Private Sub GenerarTablas()
    On Error Resume Next
    ExportarMuestra
    If Err.Number <> 0 Then
        MsgBox "No encontré 'ExportarMuestra'.", vbExclamation
        Err.Clear
    End If
End Sub

Public Sub SeleccionarMuestra()
    On Error Resume Next
    SeleccionMuestra
    If Err.Number <> 0 Then
        MsgBox "No encontré 'SeleccionMuestra'.", vbExclamation
        Err.Clear
    End If
End Sub

Public Sub CargarOrdenes_PQ()
    Dim ruta As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim connStr As String
    Dim cmdText As String

    ruta = PickExcelFilePath()
    If Len(ruta) = 0 Then Exit Sub

    On Error GoTo FIN
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.CutCopyMode = False

    ResetOrdenesEnvironment ws
    UpsertQuery "Ordenes", M_Ordenes_PQ_Spanish(ruta)

    connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Ordenes;Extended Properties=" & Chr(34) & Chr(34)
    cmdText = "SELECT * FROM [Ordenes]"

    Set lo = ws.ListObjects.Add(SourceType:=0, Source:=connStr, Destination:=ws.Range("A1"))
    With lo
        .name = "Ordenes"
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

    With EnsureSheet("Muestra")
        .Range("N3").FormulaLocal = "=CONTARA(Ordenes[NºOrden])"
        SafeDefineName "Universo", .Range("N3").Address(True, True, xlA1, True)
    End With

    On Error Resume Next
    MuestrasPorMes_Rebuild
    On Error GoTo 0

FIN:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Not ws Is Nothing And ws.ListObjects.Count > 0 Then
        MsgBox "Consulta 'Ordenes' cargada como tabla 'Ordenes' en la hoja 'Ordenes'.", vbInformation
    End If
End Sub

Private Sub ResetOrdenesEnvironment(ByRef wsOut As Worksheet)
    Dim ws As Worksheet, lo As ListObject
    Dim cN As WorkbookConnection

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Do While ws.QueryTables.Count > 0
            If InStr(1, ws.QueryTables(1).Connection, "Microsoft.Mashup.OleDb.1", vbTextCompare) > 0 _
               Or InStr(1, ws.QueryTables(1).Connection, "Location=Ordenes", vbTextCompare) > 0 _
               Or InStr(1, ws.QueryTables(1).CommandText, "[Ordenes]", vbTextCompare) > 0 Then
                ws.QueryTables(1).Delete
            Else
                Exit Do
            End If
        Loop
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects("Ordenes")
        If Not lo Is Nothing Then lo.Unlist
        On Error GoTo 0
    Next ws

    For Each cN In ThisWorkbook.Connections
        If cN.Type = xlConnectionTypeOLEDB Then
            If InStr(1, cN.OLEDBConnection.Connection, "Microsoft.Mashup.OleDb.1", vbTextCompare) > 0 _
               And InStr(1, cN.OLEDBConnection.Connection, "Location=Ordenes", vbTextCompare) > 0 Then
                cN.Delete
            End If
        End If
    Next cN

    On Error Resume Next
    ThisWorkbook.Queries("Ordenes").Delete
    On Error GoTo 0

    Set wsOut = Nothing
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Ordenes")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsOut.name = "Ordenes"
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

Private Function PickExcelFilePath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccionar archivo de Ordenes (.XLS, .CSV, .TXT)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Archivos comunes", "*.xls; *.xlsx; *.xlsm; *.xlsb; *.csv; *.txt"
        If .Show <> -1 Then Exit Function
        PickExcelFilePath = .SelectedItems(1)
    End With
End Function

' ====== M CORREGIDO (ordenado por Fecha y Hora al final) ======
Private Function M_Ordenes_PQ_Spanish(ByVal ruta As String) As String
    Dim m As String, p As String
    p = Replace(ruta, """", """""")
    m = ""
    m = m & "let" & vbCrLf
    m = m & "  path = """ & p & """," & vbCrLf & vbCrLf
    m = m & "  expected = {""NºOrden"",""Fecha"",""Hora"",""Cuenta"",""Nombre Cuenta"",""Modalidad"",""Operacion"",""Tipo Orden"",""Inst"",""Serie""," & vbCrLf
    m = m & "               ""Emisor"",""Moneda"",""Precio"",""Monto"",""Tasa"",""Plazo"",""Orden"",""Exta"",""Asig"",""Pend"",""TipOrd"",""Est""," & vbCrLf
    m = m & "               ""Observaciones"",""Oficial de Cuenta"",""VC"",""Usu Reg"",""Estado"",""Motivo"",""Email Cliente""}," & vbCrLf & vbCrLf

    m = m & "  Canon = (s as text) as text =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      t0 = Text.Upper(Text.Trim(s))," & vbCrLf
    m = m & "      t1 = Text.Replace(t0,""Á"",""A"")," & vbCrLf
    m = m & "      t2 = Text.Replace(t1,""É"",""E"")," & vbCrLf
    m = m & "      t3 = Text.Replace(t2,""Í"",""I"")," & vbCrLf
    m = m & "      t4 = Text.Replace(t3,""Ó"",""O"")," & vbCrLf
    m = m & "      t5 = Text.Replace(t4,""Ú"",""U"")," & vbCrLf
    m = m & "      t6 = Text.Replace(t5,""Ñ"",""N"")," & vbCrLf
    m = m & "      t7 = Text.Replace(Text.Replace(t6,""Nº"",""N""),""N°"",""N"")," & vbCrLf
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

    m = m & "  _fail = if TryAll = null then error ""No se encontró la fila con las 29 cabeceras en el archivo."" else null," & vbCrLf & vbCrLf

    m = m & "  tblAll  = TryAll[Tbl]," & vbCrLf
    m = m & "  hIdx    = TryAll[HeaderIdx]," & vbCrLf & vbCrLf

    m = m & "  tblAfterSkip = Table.Skip(tblAll, hIdx)," & vbCrLf
    m = m & "  promoted     = Table.PromoteHeaders(tblAfterSkip, [PromoteAllScalars=true])," & vbCrLf & vbCrLf

    m = m & "  curNames = Table.ColumnNames(promoted)," & vbCrLf
    m = m & "  pairs    = List.Transform({0..lenExp-1}, each {curNames{_}, expected{_}})," & vbCrLf
    m = m & "  renamed  = Table.RenameColumns(promoted, pairs, MissingField.Ignore)," & vbCrLf & vbCrLf

    m = m & "  only29   = Table.SelectColumns(renamed, expected, MissingField.UseNull)," & vbCrLf & vbCrLf

    m = m & "  allTextCols = List.RemoveItems(expected, {""Fecha"",""Hora""})," & vbCrLf
    m = m & "  AsText = Table.TransformColumnTypes(only29, List.Transform(allTextCols, each {_, type text}), ""es-PE"")," & vbCrLf & vbCrLf

    m = m & "  MonthPairs = {" & vbCrLf
    m = m & "    ""ene"",""enero"",""feb"",""febrero"",""mar"",""marzo"",""abr"",""abril"",""may"",""mayo"",""jun"",""junio""," & vbCrLf
    m = m & "    ""jul"",""julio"",""ago"",""agosto"",""set"",""septiembre"",""sep"",""septiembre"",""sept"",""septiembre""," & vbCrLf
    m = m & "    ""oct"",""octubre"",""nov"",""noviembre"",""dic"",""diciembre""" & vbCrLf
    m = m & "  }," & vbCrLf & vbCrLf

    m = m & "  NormalizeMonthText = (t as text) as text =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      s0 = Text.Trim(t)," & vbCrLf
    m = m & "      s1 = Text.Lower(s0)," & vbCrLf
    m = m & "      s2 = Text.Replace(Text.Replace(Text.Replace(s1, ""/"", "" ""), ""-"", "" ""), ""."", """")," & vbCrLf
    m = m & "      parts = List.Select(Text.Split(s2, "" ""), each _ <> """")," & vbCrLf
    m = m & "      s3 = Text.Combine(parts, "" "")," & vbCrLf
    m = m & "      s4 = "" "" & s3 & "" "", " & vbCrLf
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

    m = m & "  FechaFix = Table.TransformColumns(AsText, {{""Fecha"", each ParseFechaES(_), type date}})," & vbCrLf & vbCrLf

    m = m & "  HoraFix = Table.TransformColumns(" & vbCrLf
    m = m & "              FechaFix," & vbCrLf
    m = m & "              {{""Hora"", each" & vbCrLf
    m = m & "                  let x = _ in" & vbCrLf
    m = m & "                    if x is time then x" & vbCrLf
    m = m & "                    else if x is datetime then Time.From(x)" & vbCrLf
    m = m & "                    else if x is text then (try Time.FromText(x, ""es-PE"") otherwise try Time.FromText(x) otherwise null)" & vbCrLf
    m = m & "                    else null," & vbCrLf
    m = m & "                type time}}" & vbCrLf
    m = m & "            )," & vbCrLf & vbCrLf

    ' === ÚNICO CAMBIO: ordenar por Fecha y Hora ===
    m = m & "  Final = Table.TransformColumnTypes(HoraFix, {{""Precio"", type text},{""Tasa"", type text},{""Plazo"", type text}}, ""es-PE"")," & vbCrLf
    m = m & "  Sorted = Table.Sort(Final, {{""Fecha"", Order.Ascending}, {""Hora"", Order.Ascending}})" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "  Sorted" & vbCrLf

    M_Ordenes_PQ_Spanish = m
End Function

