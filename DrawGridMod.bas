Attribute VB_Name = "DrawGridMod"
Option Explicit
Option Base 1
Dim FrameOffsetX As Single
Dim FrameOffsetY As Single

Public Sub DrawGrid(Size1 As Single, NumX As Integer, NumY As Integer, SheetName As String)
    
    Dim ws1 As Worksheet: Set ws1 = Worksheets(SheetName)
    Dim n1 As Integer, m1 As Integer

    Dim x1 As Single
    Dim y1 As Single
    Dim Minor As Single: Minor = 6 'Small divisions
    Dim Size2 As Single: Size2 = Size1 / Minor
    
    Dim StartShape As Integer: StartShape = ws1.Shapes.Count
    
    'Mayor Scale
    For n1 = 0 To NumY
        x1 = FrameOffsetX
        y1 = Size1 * n1 + FrameOffsetY
        With ws1.Shapes.AddLine(x1, y1, x1 + Size1 * NumX, y1)
            .Line.Weight = 3
        End With
    Next n1
    
    For n1 = 0 To NumX
        x1 = Size1 * n1 + FrameOffsetX
        y1 = FrameOffsetY
        With ws1.Shapes.AddLine(x1, y1, x1, y1 + Size1 * NumY)
            .Line.Weight = 3
        End With
    Next n1

    'Minor Scale
    For n1 = 0 To NumY * Minor
        x1 = FrameOffsetX
        y1 = Size2 * n1 + FrameOffsetY
        With ws1.Shapes.AddLine(x1, y1, x1 + Size1 * NumX, y1)
            .Line.Weight = 1
        End With
    Next n1
    
    For n1 = 0 To NumX * Minor
        x1 = Size2 * n1 + FrameOffsetX
        y1 = FrameOffsetY
        With ws1.Shapes.AddLine(x1, y1, x1, y1 + Size1 * NumY)
            .Line.Weight = 1
        End With
    Next n1

    Dim EndShape  As Integer: EndShape = ws1.Shapes.Count
    
    Dim Sr1 As ShapeRange: Set Sr1 = ws1.Shapes.Range(NumRange(StartShape + 1, EndShape))
    Dim Grid As Shape: Set Grid = Sr1.Group
    Grid.Name = "Grid"

End Sub

Private Function NumRange(P1 As Integer, P2 As Integer) As Variant

    Dim Result As Variant
    ReDim Result(P2 - P1 + 1)
    Dim n As Integer
    For n = P1 To P2
        Result(n) = n
    Next n
    NumRange = Result
    
End Function
