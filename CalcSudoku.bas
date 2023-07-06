Attribute VB_Name = "CalcSudoku"
Option Explicit
Option Base 1

Dim Values(256) As Variant
Dim ValuesCount As Integer

Public Sub CalcSudoku()
    
    
    Values(1) = Range("b2:j10").value
    ValuesCount = 1

    Dim n As Integer
    For n = 1 To ValuesCount
    
        Dim Count As Integer: Count = 0

        Do
            If Not Iterate(n) Then Exit Do
            Count = Count + 1
            If Count > 100 Then Exit Do
            'Debug.Print Count
        Loop

    Next n


    Range("b2:j10").value = Values(1)

    ShowData

End Sub
Private Sub ShowData()

    Dim i As Integer
    Dim RowOffset As Integer: RowOffset = 12
    
    For i = 1 To ValuesCount
        Dim RowOffsetTot As Integer: RowOffsetTot = RowOffset + (i - 1) * 10
        Dim Rg1 As Range: Set Rg1 = Range(Cells(RowOffsetTot, 2), Cells(RowOffsetTot + 8, 2 + 8))
        Rg1.value = Values(i)
        Borders Rg1
    Next i

End Sub


Private Function Iterate(sol As Integer) As Boolean

   Iterate = False

   Dim m As Integer
   Dim n As Integer
   Dim v As Integer
    
   For m = 1 To 9
   For n = 1 To 9
    
    If Values(sol)(m, n) = vbNullString Then
    
        Iterate = True
    
        Dim vc As Integer: vc = 0
    
        For v = 1 To 9
            If CheckBlock(v, m, n, sol) And CheckRowCol(v, m, n, sol) Then
                If vc = 0 Then
                    Values(sol)(m, n) = v
                Else
                    ValuesCount = ValuesCount + 1
                    Values(ValuesCount) = Values(sol)
                    Values(ValuesCount)(m, n) = v
                    'Debug.Print ValuesCount
                End If
                vc = vc + 1
            End If
        Next v
        
   
    End If
   
   Next n
   Next m

   Debug.Print ValuesCount


End Function

Private Function CheckBlock(val1 As Integer, row1 As Integer, col1 As Integer, sol As Integer) As Boolean

    CheckBlock = True
        
    Dim TableRow As Integer: TableRow = ((row1 - 1) \ 3) * 3 + 1
    Dim TableCol As Integer: TableCol = ((col1 - 1) \ 3) * 3 + 1

    Dim m As Integer
    Dim n As Integer
    
    For m = TableRow To TableRow + 2
    For n = TableCol To TableCol + 2
        If Values(sol)(m, n) = val1 Then
            CheckBlock = False
        End If
    Next n
    Next m


End Function

Private Function CheckRowCol(val1 As Integer, row1 As Integer, col1 As Integer, sol As Integer) As Boolean

    CheckRowCol = True
    
    Dim i As Integer
    For i = 1 To 9
        If (Values(sol)(i, col1) = val1) Or (Values(sol)(row1, i) = val1) Then
            CheckRowCol = False
            Exit For
        End If
    Next i

End Function

Private Sub Borders(Rg1 As Range)

    With Rg1.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Rg1.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Rg1.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Rg1.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End Sub

