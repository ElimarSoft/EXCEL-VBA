Public Function RegexReplace(Data As String, Pattern As String, Values As Variant)

    'Debug.Print RegexReplace("=COUNTIF($C{Int1}:$C{Int2};F$1)", "(Int[0-9])", [{3,4}])

    Dim RegExp1 As New RegExp
    RegExp1.Pattern = Pattern
    RegExp1.Global = False
    
    Dim n As Integer
    For n = LBound(Values) To UBound(Values)
        Data = RegExp1.Replace(Data, Values(n))
    Next n
    
    RegexReplace = Data

End Function

