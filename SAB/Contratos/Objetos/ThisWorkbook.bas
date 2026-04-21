' ========== ThisWorkbook.bas ==========
Option Explicit

' ============================================================
'  Recalcula TamañoPoblacion cuando cambia la tabla Contratos.
'  Ya no monitorea Mes/Año/TipoInforme porque esos controles
'  fueron eliminados del flujo.
' ============================================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrHandler
    Application.EnableEvents = False

    If Sh.name = "Contratos" Then
        Dim lo As ListObject
        On Error Resume Next
        Set lo = Sh.ListObjects("Contratos")
        On Error GoTo 0
        If Not lo Is Nothing Then
            If Not Intersect(Target, lo.Range) Is Nothing Then
                TamañoPoblacion
            End If
        End If
    End If

ExitHandler:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    MsgBox "Error en Workbook_SheetChange: " & Err.Number & " - " & Err.Description, vbCritical
    Resume ExitHandler
End Sub