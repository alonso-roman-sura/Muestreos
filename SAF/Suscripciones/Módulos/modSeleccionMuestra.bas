' ========== modSeleccionMuestra_Suscripciones.bas ==========
Option Explicit

Public Sub SeleccionMuestra()
    Dim wb As Workbook: Set wb = ThisWorkbook

    ' Validar que haya datos importados antes de continuar
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = wb.Worksheets("Suscripciones")
    On Error GoTo 0
    If wsC Is Nothing Then
        MsgBox "No existe la hoja 'Suscripciones'." & vbCrLf & _
               "Importe los datos primero.", vbExclamation, "Sin datos"
        Exit Sub
    End If

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsC.ListObjects("Suscripciones")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        MsgBox "La tabla 'Suscripciones' est" & Chr(225) & " vac" & Chr(237) & "a." & vbCrLf & _
               "Importe los datos primero.", vbExclamation, "Sin datos"
        Exit Sub
    End If

    ' Validar que los universos estén calculados
    Dim contPN As Long, contPJ As Long
    On Error Resume Next
    contPN = CLng(wb.Names("UniversoPN").RefersToRange.Value)
    contPJ = CLng(wb.Names("UniversoPJ").RefersToRange.Value)
    On Error GoTo 0

    If contPN = 0 And contPJ = 0 Then
        MsgBox "El universo PN y PJ son ambos 0." & vbCrLf & vbCrLf & _
               "Esto puede deberse a que:" & vbCrLf & _
               "  " & Chr(149) & "  Los datos a" & Chr(250) & "n no se han importado." & vbCrLf & _
               "  " & Chr(149) & "  La columna 'TIPO PERSONA' no fue reconocida." & vbCrLf & _
               "  " & Chr(149) & "  Los nombres definidos 'UniversoPN'/'UniversoPJ' no existen.", _
               vbExclamation, "Universo vac" & Chr(237) & "o"
        Exit Sub
    End If

    Dim tamPN As Long, tamPJ As Long
    On Error Resume Next
    tamPN = CLng(wb.Names("Tama" & Chr(241) & "oMuestraPN").RefersToRange.Value)
    tamPJ = CLng(wb.Names("Tama" & Chr(241) & "oMuestraPJ").RefersToRange.Value)
    On Error GoTo 0

    If tamPN = 0 And tamPJ = 0 Then
        MsgBox "El tama" & Chr(241) & "o de muestra PN y PJ son ambos 0." & vbCrLf & vbCrLf & _
               "Verifique los par" & Chr(225) & "metros Z, p y E en la hoja Muestra.", _
               vbExclamation, "Tama" & Chr(241) & "o de muestra inv" & Chr(225) & "lido"
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox(Chr(191) & "Est" & Chr(225) & " seguro de que desea generar nuevas muestras PN y PJ?", _
                  vbYesNo + vbQuestion, "Confirmar")
    If resp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim errPN As String, errPJ As String
    errPN = GenerarMuestraOrdenada("Muestra1_PN", "Tama" & Chr(241) & "oMuestraPN", "UniversoPN")
    errPJ = GenerarMuestraOrdenada("Muestra1_PJ", "Tama" & Chr(241) & "oMuestraPJ", "UniversoPJ")

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If Len(errPN) > 0 Or Len(errPJ) > 0 Then
        Dim msgErr As String
        If Len(errPN) > 0 Then msgErr = msgErr & "PN: " & errPN & vbCrLf
        If Len(errPJ) > 0 Then msgErr = msgErr & "PJ: " & errPJ & vbCrLf
        MsgBox "Se produjeron errores al generar las muestras:" & vbCrLf & vbCrLf & msgErr, _
               vbCritical, "Error en selecci" & Chr(243) & "n"
    Else
        MsgBox "Muestras PN y PJ generadas correctamente.", vbInformation
    End If
End Sub

' Retorna string vacío si OK, o descripción del error si falla
Private Function GenerarMuestraOrdenada(nombreInicio As String, nombreTamano As String, nombreUniverso As String) As String
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim rngInicio As Range
    Dim ws As Worksheet
    Dim tamano As Long, universo As Long
    Dim coll As Collection
    Dim numeros() As Long
    Dim i As Long, j As Long, tmp As Long
    Dim startRow As Long, startCol As Long
    Dim lastRowUsed As Long
    Dim c As Long, R As Long
    Dim rnum As Long, fila As Long, col As Long

    On Error Resume Next
    Set rngInicio = wb.Names(nombreInicio).RefersToRange
    On Error GoTo 0
    If rngInicio Is Nothing Then
        GenerarMuestraOrdenada = "No existe el nombre definido '" & nombreInicio & "'."
        Exit Function
    End If
    Set ws = rngInicio.Parent

    On Error Resume Next
    tamano = CLng(wb.Names(nombreTamano).RefersToRange.Value)
    universo = CLng(wb.Names(nombreUniverso).RefersToRange.Value)
    On Error GoTo 0

    If universo = 0 Then
        GenerarMuestraOrdenada = "El universo ('" & nombreUniverso & "') es 0. No hay registros de este tipo."
        Exit Function
    End If
    If tamano = 0 Then
        GenerarMuestraOrdenada = "El tama" & Chr(241) & "o de muestra ('" & nombreTamano & "') es 0."
        Exit Function
    End If
    If tamano > universo Then
        GenerarMuestraOrdenada = "El tama" & Chr(241) & "o (" & tamano & ") supera el universo (" & universo & ")."
        Exit Function
    End If

    startRow = rngInicio.Row
    startCol = rngInicio.Column

    lastRowUsed = startRow
    For c = startCol To startCol + 4
        R = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If R > lastRowUsed Then lastRowUsed = R
    Next c

    With ws
        .Range(.Cells(startRow, startCol), .Cells(lastRowUsed, startCol + 4)).ClearContents
        If lastRowUsed > startRow Then
            .Range(.Cells(startRow + 1, startCol), .Cells(lastRowUsed, startCol + 4)).ClearFormats
        End If
        If startCol + 4 > startCol Then
            .Range(.Cells(startRow, startCol + 1), .Cells(lastRowUsed, startCol + 4)).ClearFormats
        End If
    End With

    Set coll = New Collection
    Randomize
    Do While coll.Count < tamano
        rnum = Int(universo * Rnd) + 1
        On Error Resume Next
        coll.Add rnum, CStr(rnum)
        On Error GoTo 0
    Loop

    ReDim numeros(1 To coll.Count)
    For i = 1 To coll.Count
        numeros(i) = coll(i)
    Next i
    For i = 1 To UBound(numeros) - 1
        For j = i + 1 To UBound(numeros)
            If numeros(i) > numeros(j) Then
                tmp = numeros(i): numeros(i) = numeros(j): numeros(j) = tmp
            End If
        Next j
    Next i

    fila = startRow
    col = startCol
    For i = 1 To UBound(numeros)
        ws.Cells(fila, col).Value = numeros(i)
        rngInicio.Copy
        ws.Cells(fila, col).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        With ws.Cells(fila, col).Borders
            .LineStyle = xlDot
            .Color = RGB(128, 128, 128)
            .Weight = xlHairline
        End With

        col = col + 1
        If col > startCol + 4 Then
            col = startCol
            fila = fila + 1
        End If
    Next i

    GenerarMuestraOrdenada = "" ' OK
End Function