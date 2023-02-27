Attribute VB_Name = "ChartAverageMod"
Option Explicit


Public Sub AddChart()

    'Delete previous charts if needed
    DeleteShapes
    
    Dim Shape1 As Shape: Set Shape1 = ActiveSheet.Shapes.AddChart2(227, xlLine)
    Dim Chart1 As Chart: Set Chart1 = Shape1.Chart
    
    'Set source data
    Chart1.SetSourceData Source:=Selection
    
    'Get average values
    Dim Text1 As String: Text1 = vbNullString
    Dim Ser1 As Series
    For Each Ser1 In Chart1.SeriesCollection
        Dim aver1 As Variant: aver1 = WorksheetFunction.Average(Ser1.Values)
        Text1 = Text1 + "Average = " + Format(aver1, "0.####") + vbCr
    Next
    'Fill caption with average and duplicate Chart default size
    Chart1.ChartTitle.Caption = Text1
    Chart1.ChartArea.Width = 2 * Chart1.ChartArea.Width
    Chart1.ChartArea.Height = 2 * Chart1.ChartArea.Height
    
    'Make Chart Background Invisible for comparition
    Shape1.Fill.Visible = msoFalse

End Sub

Private Sub DeleteShapes()
    
    Dim n As Integer
    For n = 1 To ActiveSheet.Shapes.Count
        ActiveSheet.Shapes(1).Delete
    Next n

End Sub


