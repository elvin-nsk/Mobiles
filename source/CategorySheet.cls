VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CategorySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Const MaxPageSize As Double = 40000 'mm
Private Const VerticalSpaceMult As Double = 0.18
Private Const HorizontalSpaceMult As Double = VerticalSpaceMult
Private Const SpaceBetweenBlocksMult As Double = 1.8
Private Const TextSizeMult As Double = 0.6

Private Type typeThis
  Category As IRecord
  AllModels As IRecordList
  AdditionalBlocks As IRecordList
  Sizes As IRecordList
  Document As Document
  FilteredModels As IRecordList
  CategoriesFromShapeSizes As Boolean
  FailedFiles As IList
  SumOfAverageSizes As Double
  ImportedModelsCount As Long
  TotalAverageSize As Double
  HorizontalSpace As Double
  VerticalSpace As Double
  FontSize As Double
  Blocks As Scripting.Dictionary
End Type
Private This As typeThis

'===============================================================================

Public Function CreateAndCompose( _
                  ByVal Category As IRecord, _
                  ByVal Models As IRecordList, _
                  ByVal AdditionalBlocks As IRecordList, _
                  ByVal Sizes As IRecordList _
                ) As IEither
  With New CategorySheet
    Set CreateAndCompose = .Init(Category, Models, AdditionalBlocks, Sizes)
  End With
End Function

Friend Function Init( _
                  ByVal Category As IRecord, _
                  ByVal Models As IRecordList, _
                  ByVal AdditionalBlocks As IRecordList, _
                  ByVal Sizes As IRecordList _
                ) As IEither
  With This
  
    Set .Category = Category
    Set .AllModels = Models
    Set .AdditionalBlocks = AdditionalBlocks
    Set .Sizes = Sizes
    
    FilterModels
    If .FilteredModels.Count = 0 Then
      Set Init = Either.Create
      Exit Function
    End If
    
    Set .Document = CreateDocument
    .Document.Name = Category!Name
    .Document.Unit = cdrMillimeter
    .Document.ActivePage.SetSize MaxPageSize, MaxPageSize
    
    Set .Blocks = New Scripting.Dictionary
    Set .FailedFiles = List.Create
    ImportAndSortModels
    
    If .Blocks.Count = 0 Then
      .Document.Close
      Set Init = Either.Create
      Exit Function
    End If
    
    CalculateSheetSizesAndSpaces
    ComposeSheet
    
    .Document.ClearSelection
    
  End With
  Set Init = Either.Create(Me)
End Function

'===============================================================================

Public Property Get Document() As Document
  Set Document = This.Document
End Property

Public Property Get FailedFiles() As IList
  Set FailedFiles = This.FailedFiles
End Property

Public Property Get Self() As CategorySheet
  Set Self = Me
End Property

'===============================================================================

Private Sub FilterModels()
  With This
    Dim Arr As Variant
    If .Category("Size") = "" Then
      Set .FilteredModels = .AllModels.Filter.FieldsEmpty("Size")
    Else
      Arr = VBA.Split(.Category("Size"), Mobiles.SizesDelimiterSymbol)
      Set .FilteredModels = .AllModels.Filter.Fields(TrimArrayAndToList(Arr), "Size")
    End If
  End With
End Sub

Private Function TrimArrayAndToList(ByRef Arr As Variant) As IList
  Set TrimArrayAndToList = List.Create
  Dim Item As Variant
  For Each Item In Arr
    TrimArrayAndToList.Add VBA.Trim$(Item)
  Next Item
End Function

Private Sub ImportAndSortModels()
  With This
    .Document.Activate
    Dim Record As IRecord
    For Each Record In .FilteredModels.NewEnum
      ImportAndSortModel Record
    Next Record
  End With
End Sub

Private Sub ImportAndSortModel(ByVal Record As IRecord)
   With This
    
    If IsFieldEmpty(Record!Count) Then Exit Sub
    
    Dim Shape As Shape
    Set Shape = ImportModel(Record!File)
    If Shape Is Nothing Then Exit Sub
    
    Shape.Name = Record!Name
    SortBySizes Record, Shape
    
    .SumOfAverageSizes = .SumOfAverageSizes + lib_elvin.AverageDim(ActiveShape)
    .ImportedModelsCount = .ImportedModelsCount + 1
    
    Shape.Delete
    
  End With
End Sub

Private Function ImportModel(ByVal File As String) As Shape
  On Error GoTo Fail
  If Not lib_elvin.FileExists(File) Then GoTo Fail
  This.Document.ActiveLayer.Import File
  Set ImportModel = ActiveShape
  Exit Function
Fail:
  On Error GoTo 0
  This.FailedFiles.Add File
End Function

Private Sub SortBySizes(ByVal Record As IRecord, ByVal Shape As Shape)
  Dim HaveSubs As Boolean
  HaveSubs = AddModelToSuitableBlockIfCountNotEmpty _
             (Record!SubCount1, Record!Size, Mobiles.SubColumn1, Shape)
  HaveSubs = HaveSubs Or AddModelToSuitableBlockIfCountNotEmpty _
                         (Record!SubCount2, Record!Size, Mobiles.SubColumn2, Shape)
  HaveSubs = HaveSubs Or AddModelToSuitableBlockIfCountNotEmpty _
                         (Record!SubCount3, Record!Size, Mobiles.SubColumn3, Shape)
  HaveSubs = HaveSubs Or AddModelToSuitableBlockIfCountNotEmpty _
                         (Record!SubCount4, Record!Size, Mobiles.SubColumn4, Shape)
  If Not HaveSubs Then AddModelToSuitableBlockIfCountNotEmpty _
                       Record!Count, Record!Size, 0, Shape
End Sub

Private Function AddModelToSuitableBlockIfCountNotEmpty _
                 (ByVal Count As String, ByVal Size As String, _
                  ByVal CountColumn As Long, ByVal Shape As Shape) As Boolean
  
  If IsFieldEmpty(Count) Then Exit Function
  
  Dim Model As New structModel
  Set Model.Shape = Shape.Duplicate
  Model.Count = VBA.CLng(Count)
  
  Dim SizeKey As String
  If Size = "" Then
    SizeKey = GetSizeFromShape(Shape)
  Else
    If CountColumn = 0 Then
      SizeKey = Size
    Else
      SizeKey = SizeKey & "_" & VBA.CStr(CountColumn)
    End If
  End If
  
  Dim Block As CategorySheetBlock
  Set Block = GetOrCreateBlock(SizeKey, CountColumn)
  Block.Add Model
  
  AddModelToSuitableBlockIfCountNotEmpty = True
  
End Function

Private Function GetSizeFromShape(ByVal Shape As Shape) As String
  GetSizeFromShape = VBA.CStr(VBA.Round(Shape.SizeWidth, 0))
  GetSizeFromShape = GetSizeFromShape & Mobiles.SizesMultiplierSymbol & _
                  VBA.CStr(VBA.Round(Shape.SizeHeight, 0))
End Function

Private Function GetOrCreateBlock(ByVal Key As String, _
                                  ByVal AdditionalBlocksID As Long _
                                  ) As CategorySheetBlock
  With This
    
    If This.Blocks.Exists(Key) Then
      Set GetOrCreateBlock = This.Blocks(Key)
      Exit Function
    End If
    
    Set GetOrCreateBlock = CategorySheetBlock.Create
    This.Blocks.Add Key, GetOrCreateBlock
        
    If AdditionalBlocksID = 0 Then
      If .Sizes.PrimaryFieldExists(Key) Then
        GetOrCreateBlock.Caption.Line1 = .Sizes(Key)("Name")
      Else
        GetOrCreateBlock.Caption.Line1 = Key
      End If
    Else
      If .AdditionalBlocks.PrimaryFieldExists(VBA.CStr(AdditionalBlocksID)) Then
        GetOrCreateBlock.Caption.Line1 = _
          .AdditionalBlocks(VBA.CStr(AdditionalBlocksID))("Name")
      Else
        GetOrCreateBlock.Caption.Line1 = Key
      End If
    End If
    
    GetOrCreateBlock.MaxPlacesInWidth = .Category!MaxPlacesInWidth
  
  End With
End Function

Private Sub CalculateSheetSizesAndSpaces()
  With This
    If .ImportedModelsCount = 0 Then Exit Sub
    .TotalAverageSize = .SumOfAverageSizes / .ImportedModelsCount
    .VerticalSpace = .TotalAverageSize * VerticalSpaceMult
    .HorizontalSpace = .TotalAverageSize * HorizontalSpaceMult
    .FontSize = .TotalAverageSize * TextSizeMult
  End With
End Sub

Private Sub ComposeSheet()
  With This
  
    Dim Shapes As New ShapeRange
    Dim Stackables As New Collection
  
    Dim Block As Variant 'CategorySheetBlock
    For Each Block In .Blocks.Items
      ComposeBlock Block
      Shapes.AddRange Block.Shapes
      Stackables.Add Stackable.Create(Block.Shapes)
    Next Block
    
    Dim StartingPoint As IPoint
    Set StartingPoint = FreePoint.Create(0, MaxPageSize)
    
    Stacker.CreateAndStack _
              Stackables:=Stackables, _
              StartingPoint:=StartingPoint, _
              MaxPlacesInWidth:=VBA.CLng(This.Category!MaxBlocksInWidth), _
              MaxWidth:=0, _
              HorizontalSpace:=This.HorizontalSpace * SpaceBetweenBlocksMult, _
              VerticalSpace:=This.VerticalSpace * SpaceBetweenBlocksMult
                           
    Dim Caption As Shape
    Set Caption = CreateCaption
    Caption.Text.Story.Alignment = cdrCenterAlignment
    Caption.CenterX = Shapes.CenterX
    Caption.BottomY = Shapes.TopY + .VerticalSpace * SpaceBetweenBlocksMult
    Shapes.Add Caption
    
    CenterOnPageAndPageResize Shapes
                           
  End With
End Sub

Private Sub ComposeBlock(ByVal Block As CategorySheetBlock)
  With This
    Block.Caption.Color = Mobiles.CaptionsColor
    Block.Caption.FontSize = .FontSize
    Block.HorizontalSpace = .HorizontalSpace
    Block.VerticalSpace = .VerticalSpace
    Block.StartingPoint.X = 0
    Block.StartingPoint.Y = MaxPageSize
    Block.Compose
  End With
End Sub

Private Function CreateCaption() As Shape
  Dim Caption As New structCaption
  With This
    Caption.Line1 = .Category!Line1
    Caption.Line2 = .Category!Line2
    Caption.Color = Mobiles.CaptionsColor
    Caption.FontSize = .FontSize
    Caption.Line1Bold = True
    Set CreateCaption = Helpers.CreateCaption(Caption)
  End With
End Function

Private Function IsFieldEmpty(ByVal Field As String) As Boolean
  IsFieldEmpty = True
  On Error GoTo ErrorHandler
  If Field = "" Then Exit Function
  If Field = "0" Then Exit Function
  If VBA.CStr(Field) = 0 Then Exit Function
ErrorHandler:
  On Error GoTo 0
  IsFieldEmpty = False
End Function

Private Sub CenterOnPageAndPageResize(ByVal Shapes As ShapeRange)
  With ActivePage
    Shapes.CenterX = .CenterX
    Shapes.All.CenterY = .CenterY
    .SizeWidth = .Shapes.All.SizeWidth + 2 * This.HorizontalSpace * SpaceBetweenBlocksMult
    .SizeHeight = .Shapes.All.SizeHeight + 2 * This.VerticalSpace * SpaceBetweenBlocksMult
  End With
End Sub
