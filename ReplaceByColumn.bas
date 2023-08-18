Attribute VB_Name = "ReplaceByColumn"
Option Explicit

Public Sub ReplaceByCol()

    'Replaces the String Str1 in the selected range by the value at Column col

    Dim Str1 As String
    Dim Col As String
    
    Str1 = "#Dir[n]"
    Col = "C"
  
    Dim Celda As Range
    
    For Each Celda In Selection
        If (Celda.Value <> vbNullString) And InStr(Celda.Value, Str1) > 0 Then
            Dim new_celda As String: new_celda = Replace(Celda.Value, Str1, vbNullString)
            Dim result As String: result = "=$" + Col + CStr(Celda.Row) + "&""" + new_celda + """"
            Celda.Value = result
        End If
   
    Next

End Sub

