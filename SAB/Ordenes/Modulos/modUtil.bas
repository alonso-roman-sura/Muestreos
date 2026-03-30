' ========== modUtil.bas ==========
Option Explicit

Public Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        sh.name = nm
    End If
    Set EnsureSheet = sh
End Function

Public Sub SafeDefineName(ByVal nm As String, ByVal refersToA1 As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add name:=nm, refersTo:="=" & refersToA1
End Sub

Public Sub ApplyLikeTemplate(tpl As Range, dst As Range, Optional ByVal mergeCells As Boolean = False)
    With dst
        If mergeCells Then .mergeCells = True
        .NumberFormat = tpl.NumberFormat
        .HorizontalAlignment = tpl.HorizontalAlignment
        .VerticalAlignment = tpl.VerticalAlignment
        .WrapText = tpl.WrapText
        .Font.name = tpl.Font.name
        .Font.Size = tpl.Font.Size
        .Font.Bold = tpl.Font.Bold
        .Font.Color = tpl.Font.Color
        .Interior.Color = tpl.Interior.Color
        Dim i As Long
        For i = 7 To 12
            With .Borders(i)
                .LineStyle = tpl.Borders(i).LineStyle
                .Weight = tpl.Borders(i).Weight
                .Color = tpl.Borders(i).Color
            End With
        Next i
    End With
End Sub

' Guarda/lee la última fila del bloque dinámico (una sola cota segura)
Public Sub StoreMuestrasEndRow(ws As Worksheet, ByVal R As Long)
    SafeDefineName "MuestrasEndRow", ws.Cells(R, "D").Address(True, True, xlA1, True)
End Sub

Public Function GetMuestrasEndRow(ws As Worksheet, ByVal defaultRow As Long) As Long
    Dim nm As name
    On Error Resume Next
    Set nm = ThisWorkbook.Names("MuestrasEndRow")
    On Error GoTo 0
    If nm Is Nothing Then
        GetMuestrasEndRow = defaultRow - 1
    Else
        GetMuestrasEndRow = nm.RefersToRange.Row
    End If
End Function

