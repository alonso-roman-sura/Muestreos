Option Explicit

Public WithEvents qt As QueryTable

Private Sub qt_AfterRefresh(ByVal Success As Boolean)
    If Not Success Then Exit Sub
    On Error Resume Next
    MuestrasPorMes_Rebuild
    On Error GoTo 0
End Sub