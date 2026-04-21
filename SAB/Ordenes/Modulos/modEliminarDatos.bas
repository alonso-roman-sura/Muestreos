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
        "   " & Chr(149) & "  Los valores generados en la hoja Muestra" & vbCrLf & _
        "   " & Chr(149) & "  Los nombres definidos dinámicos" & vbCrLf & vbCrLf & _
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
    LimpiarNombresDefinidos
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
'  Se itera de atrás hacia adelante para que el índice
'  no se desplace al ir borrando.
' ============================================================
Private Sub EliminarHojas()
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If Not EsHojaProtegida(ThisWorkbook.Worksheets(i).Name) Then
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
'  LIMPIAR NOMBRES DEFINIDOS DINÁMICOS
'  Elimina: Universo, UniversoXxx, MuestraXxx, MuestrasEndRow.
'  Conserva: InicioMuestra, Inicio muestra, y cualquier otro
'            que no empiece por los prefijos dinámicos.
'  Nombres de ámbito de hoja (contienen "!") se saltan siempre.
' ============================================================
Private Sub LimpiarNombresDefinidos()
    Dim nm As Name, sNm As String, sLow As String, i As Long
    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        sNm = nm.Name
        If InStr(sNm, "!") > 0 Then GoTo SiguienteNombre
        sLow = LCase$(sNm)
        If Left$(sLow, 8) = "universo" Then GoTo Borrar
        If Left$(sLow, 7) = "muestra" Then GoTo Borrar   ' InicioMuestra empieza por "inicio"
        If sLow = "muestrasendrow" Then GoTo Borrar
        GoTo SiguienteNombre
Borrar:
        On Error Resume Next
        nm.Delete
        On Error GoTo 0
SiguienteNombre:
    Next i
End Sub

' ============================================================
'  LIMPIAR HOJA MUESTRA
'
'  Fila 3 (plantilla): H3 y N3 pierden contenido pero conservan
'  formato, porque ApplyLikeTemplate los usa como referencia.
'  El borde inferior visible de fila 3 se guarda antes de limpiar
'  fila 4 y se repone después, ya que en Excel ese borde puede
'  estar almacenado como borde superior de la fila 4.
'
'  Filas 4+: se usa LimpiarBloqueConBordes, que avanza fila por
'  fila incluyendo:
'    - Filas con etiqueta que empiece por el prefijo esperado
'      (filas de datos dinámicos generadas por el código).
'    - Filas vacías con bordes inmediatamente siguientes
'      (filas placeholder del diseño de la plantilla).
'  Y se detiene en la primera fila vacía SIN bordes, evitando
'  borrar los parámetros u otro contenido ajeno al bloque.
' ============================================================
Private Sub LimpiarHojaMuestra()
    Dim wsM As Worksheet
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets("Muestra")
    On Error GoTo 0
    If wsM Is Nothing Then Exit Sub

    ' Fila de plantilla: solo contenido
    wsM.Range("H3").ClearContents
    wsM.Range("N3").ClearContents

    ' Guardar borde inferior de la fila plantilla antes de limpiar fila 4.
    ' En Excel ese borde puede estar almacenado como borde superior de la fila 4.
    Dim bLS_M As Long, bW_M As Long, bC_M As Long
    Dim bLS_U As Long, bW_U As Long, bC_U As Long
    On Error Resume Next
    With wsM.Cells(4, 4).Borders(xlEdgeTop)
        bLS_M = .LineStyle: bW_M = .Weight: bC_M = .Color
    End With
    With wsM.Cells(4, 10).Borders(xlEdgeTop)
        bLS_U = .LineStyle: bW_U = .Weight: bC_U = .Color
    End With
    On Error GoTo 0

    ' Bloque "Tamaño de la muestra Mes X": etiqueta D:G (cols 4-7), valor H (col 8)
    LimpiarBloqueConBordes wsM, 4, 4, 7, 8, "Tama"

    ' Bloque "Universo Mes X": etiqueta J:M (cols 10-13), valor N (col 14)
    LimpiarBloqueConBordes wsM, 4, 10, 13, 14, "Universo Mes "

    ' Reponer borde inferior de la fila plantilla (fila 3)
    If bLS_M <> xlLineStyleNone Then
        With wsM.Range("D3:H3").Borders(xlEdgeBottom)
            .LineStyle = bLS_M: .Weight = bW_M: .Color = bC_M
        End With
    End If
    If bLS_U <> xlLineStyleNone Then
        With wsM.Range("J3:N3").Borders(xlEdgeBottom)
            .LineStyle = bLS_U: .Weight = bW_U: .Color = bC_U
        End With
    End If

    ' Bloque de números de muestra (desde InicioMuestra)
    LimpiarBloqueInicioMuestra wsM
End Sub

' ============================================================
'  LIMPIEZA INTELIGENTE DEL BLOQUE DINÁMICO
'
'  Avanza desde startRow fila por fila:
'    - Incluye la fila si su celda de etiqueta (colLblIni) empieza
'      por el prefijo esperado ? fila de datos dinámica.
'    - Incluye la fila si está vacía pero tiene bordes en el rango
'      de etiqueta ? placeholder del diseño de la plantilla.
'    - Para en la primera fila vacía SIN bordes ? no toca
'      parámetros ni otro contenido fuera del bloque.
'    - Para si encuentra contenido que no coincida con el prefijo.
'  Al identificar el bloque completo, lo limpia de una vez con
'  .Clear (contenido + formato) previa desmerge.
' ============================================================
Private Sub LimpiarBloqueConBordes(ws As Worksheet, ByVal startRow As Long, _
                                    ByVal colLblIni As Long, ByVal colLblFin As Long, _
                                    ByVal colVal As Long, ByVal prefix As String)
    Dim R As Long, v As Variant, lastBlockRow As Long
    lastBlockRow = startRow - 1
    R = startRow

    Do While R <= startRow + 500   ' tope de seguridad
        v = ws.Cells(R, colLblIni).Value
        If Not IsEmpty(v) And Len(CStr(v)) > 0 Then
            If Left$(CStr(v), Len(prefix)) = prefix Then
                lastBlockRow = R
            Else
                Exit Do   ' contenido ajeno al bloque
            End If
        Else
            ' Fila vacía: incluir solo si tiene bordes visibles
            If FilaTieneBordes(ws, R, colLblIni, colLblFin) Then
                lastBlockRow = R
            Else
                Exit Do   ' vacía y sin bordes: fin del bloque
            End If
        End If
        R = R + 1
    Loop

    If lastBlockRow < startRow Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, colLblIni), ws.Cells(lastBlockRow, colVal))
    On Error Resume Next
    rng.UnMerge
    On Error GoTo 0
    rng.Clear
End Sub

' Devuelve True si la fila R tiene algún borde visible en el rango colIni:colFin.
Private Function FilaTieneBordes(ws As Worksheet, ByVal R As Long, _
                                  ByVal colIni As Long, ByVal colFin As Long) As Boolean
    Dim c As Long
    On Error Resume Next
    For c = colIni To colFin
        If ws.Cells(R, c).Borders(xlEdgeTop).LineStyle <> xlLineStyleNone Then
            FilaTieneBordes = True: Exit Function
        End If
        If ws.Cells(R, c).Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then
            FilaTieneBordes = True: Exit Function
        End If
        If ws.Cells(R, c).Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone Then
            FilaTieneBordes = True: Exit Function
        End If
        If ws.Cells(R, c).Borders(xlEdgeRight).LineStyle <> xlLineStyleNone Then
            FilaTieneBordes = True: Exit Function
        End If
    Next c
    On Error GoTo 0
    FilaTieneBordes = False
End Function

' ============================================================
'  LIMPIAR BLOQUE DE NÚMEROS DE MUESTRA (InicioMuestra)
'
'  Cuenta cuántos bloques de mes existen escaneando la fila de
'  títulos para acotar el ancho exacto, y borra el rango completo
'  con .Clear. Así no toca celdas adyacentes fuera del bloque.
' ============================================================
Private Sub LimpiarBloqueInicioMuestra(ws As Worksheet)
    Dim nm As Name
    On Error Resume Next
    Set nm = ThisWorkbook.Names("InicioMuestra")
    If nm Is Nothing Then Set nm = ThisWorkbook.Names("Inicio muestra")
    On Error GoTo 0
    If nm Is Nothing Then Exit Sub

    Dim inicio As Range
    On Error Resume Next
    Set inicio = nm.RefersToRange
    On Error GoTo 0
    If inicio Is Nothing Then Exit Sub

    ' Contar bloques: cada título empieza por "Muestra" y ocupa 6 columnas
    Dim nBloques As Long, offset As Long, v As Variant
    nBloques = 0
    offset = 0
    Do
        v = inicio.offset(0, offset).Value
        If IsEmpty(v) Or Len(CStr(v)) = 0 Then Exit Do
        If Left$(CStr(v), 7) <> "Muestra" Then Exit Do
        nBloques = nBloques + 1
        offset = offset + 6
    Loop

    If nBloques = 0 Then Exit Sub

    ' Última columna usada: posición 4 (0-based) dentro del último bloque
    Dim lastCol As Long
    lastCol = inicio.Column + (nBloques - 1) * 6 + 4

    ' Última fila con datos en la columna de inicio
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, inicio.Column).End(xlUp).Row
    If lastRow < inicio.Row Then lastRow = inicio.Row

    Dim rng As Range
    Set rng = ws.Range(inicio, ws.Cells(lastRow, lastCol))
    On Error Resume Next
    rng.UnMerge
    On Error GoTo 0
    rng.Clear
End Sub