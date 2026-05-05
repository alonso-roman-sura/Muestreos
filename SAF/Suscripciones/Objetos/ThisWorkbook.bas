' ========== ThisWorkbook.bas ==========
Option Explicit

' ============================================================
'  Recalcula TamañoPoblacion cuando:
'  1) Se modifica la tabla Operaciones directamente.
'  2) El usuario cambia los controles Mes, Año o TipoInforme.
' ============================================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrHandler
    Application.EnableEvents = False

    ' 1) Cambio en la tabla Operaciones
    If Sh.name = "Operaciones" Then
        Dim lo As ListObject
        On Error Resume Next
        Set lo = Sh.ListObjects("Operaciones")
        On Error GoTo 0
        If Not lo Is Nothing Then
            If Not Intersect(Target, lo.Range) Is Nothing Then
                TamañoPoblacion
                GoTo ExitHandler
            End If
        End If
    End If

    ' 2) Cambio en controles de período en la hoja Muestra
    If Sh.name = "Muestra" Then
        Dim wb As Workbook: Set wb = ThisWorkbook
        Dim nmNames As Variant
        nmNames = Array("Mes", "A" & Chr(241) & "o", "TipoInforme")
        Dim i As Long, nmRange As Range
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