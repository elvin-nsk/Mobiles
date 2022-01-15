VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Block"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Models As Collection
  Caption As structCaption
  MaxPlacesInWidth As Long
  MaxPlacesInHeight As Long
  HorizontalSpace As Double
  VerticalSpace As Double
  Cursor As IPoint
  Shapes As ShapeRange
End Type
Private This As typeThis

'===============================================================================

Public Function Create( _
                  ByVal Models As Collection, _
                  ByVal Caption As structCaption, _
                  ByVal MaxPlacesInWidth As Long, _
                  ByVal MaxPlacesInHeight As Long, _
                  ByVal HorizontalSpace As Double, _
                  ByVal VerticalSpace As Double, _
                  ByVal StartingPoint As IPoint _
                ) As Block
  With New Block
    Set Create = .Init(Models, Caption, MaxPlacesInWidth, MaxPlacesInHeight, _
                       HorizontalSpace, VerticalSpace, StartingPoint)
  End With
End Function

Friend Function Init( _
                  ByVal Models As Collection, _
                  ByVal Caption As structCaption, _
                  ByVal MaxPlacesInWidth As Long, _
                  ByVal MaxPlacesInHeight As Long, _
                  ByVal HorizontalSpace As Double, _
                  ByVal VerticalSpace As Double, _
                  ByVal StartingPoint As IPoint _
                ) As Block
  With This
    Set .Models = Models
    Set .Caption = Caption
    .MaxPlacesInWidth = MaxPlacesInWidth
    .MaxPlacesInHeight = MaxPlacesInHeight
    .VerticalSpace = VerticalSpace
    .HorizontalSpace = HorizontalSpace
    Set .Cursor = StartingPoint.GetCopy
    Set .Shapes = New ShapeRange
  End With
  Set Init = Me
End Function

'===============================================================================

Public Property Get Shapes() As ShapeRange
  Set Shapes = This.Shapes
End Property

Public Property Get ModelsCount() As Long
  ModelsCount = This.Models.Count
End Property

Public Property Get PlacesCount() As Long
  'PlacesCount = This.Models.Count
End Property

Public Property Get Self() As Block
  Set Self = Me
End Property

'===============================================================================

Private Sub Align()
  With This
    Dim CurrentRightEdge As Double
    Dim Column As structMobilesColumn
    For Each Column In .Columns
      AlignColumnAndAddCaption Column
      Column.Shapes.LeftX = CurrentRightEdge + .HorizontalSpace
      'сдвигаем, чтобы удвоить ограничение на ширину страницы
      .Document.ActivePage.Shapes.All.CenterX = 0
      Column.CaptionShape.CenterX = Column.Shapes.CenterX
      CurrentRightEdge = Column.Shapes.RightX
    Next Column
  End With
End Sub
