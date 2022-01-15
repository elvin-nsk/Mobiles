VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FreePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IPoint

'===============================================================================

Private Type typeThis
  X As Double
  Y As Double
  Frozen As Boolean
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal X As Double, ByVal Y As Double) As IPoint
  With New FreePoint
    Set Create = .Init(X, Y)
  End With
End Function

Friend Function Init(ByVal X As Double, ByVal Y As Double) As IPoint
  With This
    .X = X
    .Y = Y
  End With
  Set Init = Me
End Function

'===============================================================================

Private Function IPoint_GetCopy() As IPoint
  Set IPoint_GetCopy = Create(This.X, This.Y)
End Function

Private Function IPoint_GetDistanceFrom(ByVal Point As IPoint) As Double
End Function

Private Sub IPoint_Move(ByVal DeltaX As Double, ByVal DeltaY As Double)
  With This
    .X = .X + DeltaX
    .Y = .Y + DeltaY
  End With
End Sub

Private Sub IPoint_RotateAroundPoint(ByVal RotationCenter As IPoint, ByVal Angle As Double)
  With RotationCenter
    This.X = (This.X - .X) * VBA.Cos(Math.DegToRad(Angle)) - _
             (This.Y - .Y) * VBA.Sin(Math.DegToRad(Angle)) + .X
    This.Y = (This.X - .X) * VBA.Sin(Math.DegToRad(Angle)) + _
             (This.Y - .Y) * VBA.Cos(Math.DegToRad(Angle)) + .Y
  End With
End Sub

Private Property Get IPoint_Self() As IPoint
  Set IPoint_Self = Me
End Property

Private Property Get IPoint_X() As Double
  IPoint_X = This.X
End Property
Private Property Let IPoint_X(ByVal Value As Double)
  This.X = Value
End Property

Private Property Get IPoint_Y() As Double
  IPoint_Y = This.Y
End Property
Private Property Let IPoint_Y(ByVal Value As Double)
  This.Y = Value
End Property

'===============================================================================

Private Property Get Self() As IPoint
  Set Self = Me
End Property
