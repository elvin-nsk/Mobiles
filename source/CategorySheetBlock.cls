VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CategorySheetBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Models As IList
  Caption As structCaption
  MaxPlacesInWidth As Long
  HorizontalSpace As Double
  VerticalSpace As Double
  StartingPoint As IPoint
  Shapes As ShapeRange
End Type
Private This As typeThis

'===============================================================================

Public Function Create() As CategorySheetBlock
  With New CategorySheetBlock
    Set Create = .Init()
  End With
End Function

Friend Function Init() As CategorySheetBlock
  With This
    Set .Models = List.Create
    Set .Caption = New structCaption
    Set .StartingPoint = FreePoint.Create(0, 0)
  End With
  Set Init = Me
End Function

'===============================================================================

Public Sub Add(ByVal Model As structModel)
  This.Models.Add Model
End Sub

Public Property Get Caption() As structCaption
  Set Caption = This.Caption
End Property
Public Property Set Caption(ByVal RHS As structCaption)
  Set This.Caption = RHS
End Property

Public Property Get MaxPlacesInWidth() As Long
  MaxPlacesInWidth = This.Caption
End Property
Public Property Let MaxPlacesInWidth(ByVal RHS As Long)
  This.MaxPlacesInWidth = RHS
End Property

Public Property Get HorizontalSpace() As Double
  HorizontalSpace = This.Caption
End Property
Public Property Let HorizontalSpace(ByVal RHS As Double)
  This.HorizontalSpace = RHS
End Property

Public Property Get VerticalSpace() As Double
  VerticalSpace = This.Caption
End Property
Public Property Let VerticalSpace(ByVal RHS As Double)
  This.VerticalSpace = RHS
End Property

Public Property Get StartingPoint() As IPoint
  Set StartingPoint = This.StartingPoint
End Property
Public Property Set StartingPoint(ByVal RHS As IPoint)
  Set This.StartingPoint = RHS
End Property

Public Property Get Shapes() As ShapeRange
  Set Shapes = This.Shapes
End Property

Public Property Get ModelsCount() As Long
  ModelsCount = This.Models.Count
End Property

Public Property Get PlacesCount() As Long
  Dim Model As structModel
  For Each Model In This.Models
    PlacesCount = PlacesCount + Model.Count
  Next Model
End Property

Public Sub Compose()
  With This
  
    Dim CaptionShape As Shape
    Set CaptionShape = Helpers.CreateCaption(.Caption)
    CaptionShape.LeftX = .StartingPoint.X
    CaptionShape.TopY = .StartingPoint.Y
    .StartingPoint.Y = CaptionShape.BottomY - .VerticalSpace
      
    FillShapes
    Stacker.CreateAndStack Stackables:=GetStackables, _
                           StartingPoint:=.StartingPoint, _
                           MaxPlacesInWidth:=.MaxPlacesInWidth, _
                           MaxWidth:=0, _
                           HorizontalSpace:=.HorizontalSpace, _
                           VerticalSpace:=.VerticalSpace
    
    Shapes.Add CaptionShape
    
  End With
End Sub

Public Property Get Self() As CategorySheetBlock
  Set Self = Me
End Property

'===============================================================================

Private Function GetStackables() As Collection
  Set GetStackables = New Collection
  Dim Shape As Shape
  For Each Shape In This.Shapes
    GetStackables.Add Stackable.Create(Shape)
  Next Shape
End Function

Private Sub FillShapes()
  With This
    Set .Shapes = New ShapeRange
    Dim Model As structModel
    For Each Model In .Models.NewEnum
      .Shapes.Add Model.Shape
      If Model.Count > 1 Then
        Dim i As Long
        For i = 1 To Model.Count - 1
          .Shapes.Add Model.Shape.Duplicate
        Next
      End If
    Next Model
  End With
End Sub
