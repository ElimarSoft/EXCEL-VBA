'Select a Row double-clicking any cell

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Rows(Target.Row).Select
End Sub
