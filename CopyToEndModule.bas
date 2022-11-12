Attribute VB_Name = "CopyToEndModule"
Option Explicit

'Select a cell to copy the column down to the end of sheet
'Skip blank or boldfaced rows 

Public Sub CopyToEnd()

    Dim RowNum As Integer: RowNum = ActiveCell.Row
    Dim ColNum As Integer: ColNum = ActiveCell.Column
        
    Dim RowLast As Integer: RowLast = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
       
    Selection.Copy
    Dim n As Integer
    
    For n = RowNum + 1 To RowLast
        If (Cells(n, ColNum - 1) <> vbNullString And Not Cells(n, 1).Font.Bold) Then
            Cells(n, ColNum).PasteSpecial Paste:=xlPasteValues
        End If
    Next n
    
    
End Sub


