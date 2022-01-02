VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompositeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Const VerticalSpaceMult As Double = 0.1
Private Const HorizontalSpaceMult As Double = 0.5
Private Const TextSizeMult As Double = 0.75

Private Type typeThis
  Mobiles As Dictionary
  Boost As Boolean
  Document As Document
  Columns As Collection
  SumOfAverageSizes As Double
  ImportedMobilesCount As Long
  TotalAverageSize As Double
  VerticalSpace As Double
  HorizontalSpace As Double
  FailedFiles As Collection
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Mobiles As Dictionary, _
                       Optional ByVal Boost As Boolean = True _
                       ) As CompositeSheet
  With New CompositeSheet
    Set Create = .Init(Mobiles, Boost)
  End With
End Function

Friend Function Init(ByVal Mobiles As Dictionary, _
                     ByVal Boost As Boolean _
                     ) As CompositeSheet
  With This
    Set .Mobiles = Mobiles
    .Boost = Boost
    Set .Columns = New Collection
    Set .FailedFiles = New Collection
    Set .Document = CreateDocument
    .Document.Name = "Навигационные мобайлы"
    .Document.Unit = cdrMillimeter
    .Document.PreserveSelection = Not .Boost
    ImportAndFillColumns
    AlignColumns
    CenterOnPageAndPageResize
    .Document.PreserveSelection = .Boost
    .Document.ClearSelection
  End With
  Set Init = Me
End Function

'===============================================================================

Public Property Get Document() As Document
  Set Document = This.Document
End Property

Public Property Get FailedFiles() As Collection
  Set FailedFiles = This.FailedFiles
End Property

Public Property Get Self() As CompositeSheet
  Set Self = Me
End Property

'===============================================================================

Private Sub ImportAndFillColumns()
  With This
    .Document.Activate
    Dim PBar As IProgressBar
    Set PBar = ProgressBar.CreateNumeric(.Mobiles.Count)
    PBar.Caption = "Импорт мобайлов..."
    Dim Column As structMobilesColumn
    Dim Mobile As structMobile
    Dim Key As Variant
    For Each Key In .Mobiles.Keys
      ImportToColumns .Mobiles.Item(Key)
      PBar.Update
    Next Key
    If .ImportedMobilesCount > 0 Then _
      .TotalAverageSize = .SumOfAverageSizes / .ImportedMobilesCount
    .VerticalSpace = .TotalAverageSize * VerticalSpaceMult
    .HorizontalSpace = .TotalAverageSize * HorizontalSpaceMult
  End With
End Sub

Private Sub ImportToColumns(ByVal Mobile As structMobile)
  With This
    Dim File As String
    Dim MobilesColumn As structMobilesColumn
    File = Mobile.File.ToString
    If Not lib_elvin.FileExists(File) Then
      .FailedFiles.Add File
      Exit Sub
    End If
    .Document.ActiveLayer.Import File
    Set MobilesColumn = GetColumnByCount(Mobile.Count)
    MobilesColumn.Shapes.Add ActiveShape
    .SumOfAverageSizes = .SumOfAverageSizes + lib_elvin.AverageDim(ActiveShape)
    .ImportedMobilesCount = .ImportedMobilesCount + 1
  End With
End Sub

Private Function GetColumnByCount(ByVal Count As Long) As structMobilesColumn
  Dim MobilesColumn As structMobilesColumn
  For Each MobilesColumn In This.Columns
    If MobilesColumn.Count = Count Then
      Set GetColumnByCount = MobilesColumn
      Exit Function
    End If
  Next MobilesColumn
  Set GetColumnByCount = New structMobilesColumn
  Set GetColumnByCount.Shapes = CreateShapeRange
  GetColumnByCount.Count = Count
  This.Columns.Add GetColumnByCount
End Function

Private Sub AlignColumns()
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

Private Sub AlignColumnAndAddCaption(ByVal Column As structMobilesColumn)
  With Column
    Set .CaptionShape = ActiveLayer.CreateArtisticText _
                        (0, .Shapes.TopY, "по " & .Count & " шт.", , , , _
                         TextSizeMult * This.TotalAverageSize, , _
                         cdrTrue, , cdrCenterAlignment)
    Dim CurrentBottomEdge As Double
    CurrentBottomEdge = This.VerticalSpace
    Dim Shape As Shape
    For Each Shape In .Shapes
      Shape.TopY = CurrentBottomEdge
      CurrentBottomEdge = Shape.BottomY - This.VerticalSpace
    Next Shape
  End With
End Sub

Private Sub CenterOnPageAndPageResize()
  With ActivePage
    .Shapes.All.CenterX = .CenterX
    .Shapes.All.CenterY = .CenterY
    .SizeWidth = .Shapes.All.SizeWidth + This.HorizontalSpace
    .SizeHeight = .Shapes.All.SizeHeight + This.HorizontalSpace
  End With
End Sub
