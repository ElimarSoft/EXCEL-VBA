Attribute VB_Name = "RawToTableMod"
Option Explicit

Public Sub Test()
    RawToTable 4, True
End Sub

Private Sub RawToTable(newCols As Integer, transpose As Boolean)

    Dim Var1 As Variant
    Dim Var2 As Variant
    
    Dim row1 As Integer: row1 = Selection.row
    Dim col1 As Integer: col1 = Selection.Column
    
    Dim DataCount As Integer: DataCount = Selection.Rows.Count
    ReDim Var1(DataCount, 1)
    
    Var1 = Range(Cells(row1, col1), Cells(row1 + DataCount, col1)).Value
    
    Dim newRows As Integer: newRows = DataCount / newCols
    Dim n As Integer, m As Integer
    Dim cnt As Integer: cnt = 1
    ReDim Var2(newRows, newCols)
    
    
    If (transpose) Then
        For n = 1 To newCols
            For m = 1 To newRows
                Var2(m, n) = Var1(cnt, 1)
                cnt = cnt + 1
            Next m
        Next n
    Else
        For n = 1 To newRows
            For m = 1 To newCols
                Var2(n, m) = Var1(cnt, 1)
                cnt = cnt + 1
            Next m
        Next n
    End If
    
    Range(Cells(row1, col1 + 1), Cells(row1 + newRows, col1 + 1 + newCols)).Value = Var2


End Sub
