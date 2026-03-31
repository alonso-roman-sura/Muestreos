' ========== ThisWorkbook.bas ==========
Option Explicit

' ============================================================
'  Recalcula TamañoPoblacion cuando cambia la tabla Contratos
'  o cuando cambia alguna celda de control en la hoja Muestra.
'  CargarDatos fue movido a modPQ_Contratos.CargarContratos_PQ.
' ============================================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrHandler
    Application.EnableEvents = False

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim lo As ListObject
    Dim nm As Name, nmRange As Range, i As Long
    Dim nmNames As Variant

    ' 1) Cambio en la tabla Contratos → recalcular
    If Sh.Name = "Contratos" Then
        On Error Resume Next
        Set lo = Sh.ListObjects("Contratos")
        On Error GoTo 0
        If Not lo Is Nothing Then
            If Not Intersect(Target, lo.Range) Is Nothing Then
                TamañoPoblacion
                GoTo ExitHandler
            End If
        End If
    End If

    ' 2) Cambio en celdas de control de la hoja Muestra → recalcular
    If Sh.Name = "Muestra" Then
        nmNames = Array("Mes", "Año", "TipoInforme")
        For i = LBound(nmNames) To UBound(nmNames)
            Set nmRange = Nothing
            On Error Resume Next
            Set nmRange = wb.Names(nmNames(i)).RefersToRange
            On Error GoTo 0
            If Not nmRange Is Nothing Then
                If Not Intersect(Target, nmRange) Is Nothing Then
                    TamañoPoblacion
                    Exit For
                End If
            End If
        Next i
    End If

ExitHandler:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    MsgBox "Error en Workbook_SheetChange: " & Err.Number & " - " & Err.Description, vbCritical
    Resume ExitHandler
End Sub
