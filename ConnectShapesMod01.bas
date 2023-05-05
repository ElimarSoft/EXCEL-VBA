Attribute VB_Name = "ConnectShapesMod01"
Option Explicit

Dim ws1 As Worksheet

Enum Pos
    Base
    Up
    Left
    Down
    Right
End Enum

Public Sub Process()

    Set ws1 = Worksheets("Sheet1")
   
    DeleteShapes
    
    Dim Cep1 As Shape: Set Cep1 = DrawCep(100, 200, 300, 10)
    Dim Cep2 As Shape: Set Cep2 = DrawCep(500, 300, 100, 300)
    
    Dim Cajas(24) As Shape
    Dim n As Integer
    
    'horizontal
    For n = 0 To UBound(Cajas)
        Set Cajas(n) = ws1.Shapes.AddShape(msoShapeRoundedRectangle, 100 + n * 5, 300, 3, 3)
        DrawConnector Cep1, Cajas(n), n + 1
    Next n
    
    'vertical
'    For n = 0 To UBound(Cajas)
'        Set Cajas(n) = ws1.Shapes.AddShape(msoShapeRoundedRectangle, 100, 300 + n * 5, 3, 3)
'        DrawConnector Cep2, Cajas(n), n + 1
'    Next n

End Sub
Private Function DrawConnector(CEP As Shape, Caja As Shape, posCEP As Integer)

    Dim PosNum As Integer
    Dim cn1 As Shape
    Set cn1 = ws1.Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, 100, 400, 400, 400) 'Coordinates useless

    
    If CEP.Width > CEP.Height Then
        
        PosNum = Pos.Up
        If CEP.Top > Caja.Top Then PosNum = Pos.Down
        
        Call cn1.ConnectorFormat.BeginConnect(CEP, posCEP + 5)
        Call cn1.ConnectorFormat.EndConnect(Caja, PosNum)

    Else
    
        PosNum = Pos.Left
        If CEP.Left > Caja.Left Then PosNum = Pos.Right
        
        Call cn1.ConnectorFormat.BeginConnect(CEP, posCEP + 5)
        Call cn1.ConnectorFormat.EndConnect(Caja, PosNum)
    
    End If

'    Dim n1 As Integer
'    For n1 = 1 To cn1.Adjustments.Count
'        Debug.Print cn1.Adjustments(n1)
'    Next

    cn1.ShapeStyle = msoLineStylePreset10
    'cn1.Line.BeginArrowheadStyle = msoArrowheadOval
    With cn1.Line
        .EndArrowheadStyle = msoArrowheadTriangle
        .EndArrowheadWidth = msoArrowheadNarrow
        .EndArrowheadLength = msoArrowheadShort
        .BeginArrowheadStyle = msoArrowheadNone
    End With
    Set DrawConnector = cn1

End Function

Private Function DrawCep(x1 As Single, y1 As Single, w1 As Single, h1 As Single) As Shape
    
  Dim sh1 As Shape
  Dim PosNum As Integer
  Dim n As Integer
  
  'Set sh1 = ws1.Shapes.AddShape(msoFreeform, x1, y1, w1, h1)
  Set sh1 = ws1.Shapes.AddShape(msoShapeRectangle, x1, y1, w1, h1)
            
            
  If w1 > h1 Then
            
    PosNum = w1 / 5
          
    For n = 1 To PosNum
      Call sh1.Nodes.Insert(n + 4, msoSegmentLine, msoEditingAuto, x1 + n * 5, y1)
    Next n
        
  Else
  
    PosNum = h1 / 5
          
    For n = 1 To PosNum
      Call sh1.Nodes.Insert(n + 4, msoSegmentLine, msoEditingAuto, x1, y1 + n * 5)
    Next n
  
  End If
        
  Set DrawCep = sh1

End Function


Private Sub DeleteShapes()

    Dim sh1 As Shape
    For Each sh1 In ws1.Shapes
        ws1.Shapes(1).Delete
    Next
End Sub

