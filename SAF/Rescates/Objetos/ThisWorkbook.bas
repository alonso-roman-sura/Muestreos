Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    On Error GoTo ErrHandler
    Application.EnableEvents = False

    If sh.name = "Rescates" Then
        Dim lo As ListObject
        On Error Resume Next
        Set lo = sh.ListObjects("Rescates")
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
