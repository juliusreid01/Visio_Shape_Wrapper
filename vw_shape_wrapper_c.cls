VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_shape_wrapper_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' wrapper for visio shapes to reduce code line length
Private s As Shape

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
  Set s = Nothing
End Sub

Public Property Set vsoShape(vsoShp as Shape)
  Set s = vsoShp
End Property

Public Sub Delete()
  s.Delete
  Set s = Nothing
End Sub

Public Property Get Name as String
  Name = s.Name
End Property

Public Property Let Name(v as String)
  s.Name = v
End Property

Public Property Get Width as Variant
  Width = s.Cells("Width").Result("")
End Property

Public Property Let Width(v as Variant)
  s.Cells("Width").FormulaForce = v
End Property

Public Property Get Height as Variant
  Height = s.Cells("Height").Result("")
End Property

Public Property Let Height(v as Variant)
  s.Cells("Height").FormulaForce = v
End Property

Public Property Get Angle as Variant
  Angle = s.Cells("Angle").Result("")
End Property

Public Property Let Angle(v as Variant)
  s.Cells("Angle").FormulaForce = v
End Property

Public Property Get PinX as Variant
  PinX = s.Cells("PinX").Result("")
End Property

Public Property Let PinX(v as Variant)
  s.Cells("PinX").FormulaForce = v
End Property

Public Property Get PinY as Variant
  PinY = s.Cells("PinY").Result("")
End Property

Public Property Let PinY(v as Variant)
  s.Cells("PinY").FormulaForce = v
End Property

Public Property Get LocPinX as Variant
  LocPinX = s.Cells("LocPinX").Result("")
End Property

Public Property Let LocPinX(v as Variant)
  s.Cells("LocPinX").FormulaForce = v
End Property

Public Property Get LocPinY as Variant
  LocPinY = s.Cells("LocPinY").Result("")
End Property

Public Property Let LocPinY(v as Variant)
  s.Cells("LocPinY").FormulaForce = v
End Property

Public Property Get FlipY as Variant
  FlipY = s.Cells("FlipY").Result("")
End Property

Public Property Let FlipY(v as Variant)
  s.Cells("FlipY").FormulaForce = v
End Property

' 1-D Shape Cells
Public Property Get BeginX as Variant
  If s.CellExists("BeginX", visExistsLocally) = False Then
    BeginX = PinX - LocPinX
  Else
    BeginX = s.Cells("BeginX").Result("")
  End If
End Property

Public Property Let BeginX(v as Variant)
  If s.CellExists("BeginX", visExistsLocally) = False Then
    s.Cells("PinX").FormulaForce = v & "+" & LocPinX
  Else
    s.Cells("BeginX").FormulaForce = v
  End If
End Property

Public Property Get BeginY as Variant
  If s.CellExists("BeginY", visExistsLocally) = False Then
    BeginY = PinY - LocPinY
  Else
    BeginY = s.Cells("BeginY").Result("")
  End If
End Property

Public Property Let BeginY(v as Variant)
  If s.CellExists("BeginY", visExistsLocally) = False Then
    s.Cells("PinY").FormulaForce = v & "+" & LocPinY
  Else
    s.Cells("BeginY").FormulaForce = v
  End If
End Property

Public Property Get EndX as Variant
  If s.CellExists("EndX", visExistsLocally) = False Then
    EndX = PinX + LocPinX
  Else
    EndX = s.Cells("EndX").Result("")
  End If
End Property

Public Property Let EndX(v as Variant)
  If s.CellExists("EndX", visExistsLocally) = False Then
    s.Cells("PinX").FormulaForce = v & "-" & LocPinX
  Else
    s.Cells("EndX").FormulaForce = v
  End If
End Property

Public Property Get EndY as Variant
  If s.CellExists("EndY", visExistsLocally) = False Then
    EndY = PinY + LocPinY
  Else
    EndY = s.Cells("EndY").Result("")
  End If
End Property

Public Property Let EndY(v as Variant)
  If s.CellExists("EndY", visExistsLocally) = False Then
    s.Cells("PinY").FormulaForce = v & "-" & LocPinY
  Else
    s.Cells("EndY").FormulaForce = v
  End If
End Property

Public Function GetCellRefName(CellName as String) as String
  GetCellRefName = s.Name & "!" & CellName
End Function

Public Sub SetPoint(Section as Integer, Row as Integer, X as Variant, Y as Variant)
  s.CellsSRC(Section, Row, visX).FormulaForce = X
  s.CellsSRC(Section, Row, visY).FormulaForce = Y
End Sub
