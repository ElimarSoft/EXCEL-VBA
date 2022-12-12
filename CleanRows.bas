Attribute VB_Name = "CleanRows"
Option Explicit

Public Sub CleanRows(rowNum As Integer, colNum As Integer)
    Dim row1 As Range
    Dim data As String
    Dim ws1 As Worksheet
    
    Set ws1 = ActiveSheet
    data = ws1.Cells(rowNum, colNum)
    
    For Each row1 In ws1.UsedRange.Rows
        If (ws1.Cells(row1.Row, colNum) = data) Then row1.Delete
    Next

End Sub
