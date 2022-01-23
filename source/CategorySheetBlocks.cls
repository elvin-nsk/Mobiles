VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CategorySheetBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Category As IRecord
  AdditionalBlocks As IRecordList
  Sizes As IRecordList
  SortingStrategy As IGroupingStrategy
  Blocks As Scripting.Dictionary
End Type
Private This As typeThis

'===============================================================================

Public Function Create( _
                  ByVal Category As IRecord, _
                  ByVal AdditionalBlocks As IRecordList, _
                  ByVal Sizes As IRecordList, _
                  ByVal SortingStrategy As IGroupingStrategy _
                ) As CategorySheetBlocks
  With New CategorySheetBlocks
    Set Create = .Init(Category, AdditionalBlocks, Sizes, SortingStrategy)
  End With
End Function

Friend Function Init( _
                  ByVal Category As IRecord, _
                  ByVal AdditionalBlocks As IRecordList, _
                  ByVal Sizes As IRecordList, _
                  ByVal SortingStrategy As IGroupingStrategy _
                ) As CategorySheetBlocks
  With This
    Set .Category = Category
    Set .AdditionalBlocks = AdditionalBlocks
    Set .Sizes = Sizes
    Set .SortingStrategy = SortingStrategy
    Set .Blocks = New Scripting.Dictionary
  End With
  Set Init = Me
End Function

'===============================================================================

Public Sub AddAndSort(ByVal ModelRecord As IRecord, ByVal Shape As Shape)
  SortBySizes ModelRecord, Shape
End Sub

Public Property Get Block(ByVal Name As String) As CategorySheetBlock
  Set Block = This.Blocks.Item(Name)
End Property

Public Property Get Count() As Long
  Count = This.Blocks.Count
End Property

Public Function Exists(ByVal Name As String) As Boolean
  Exists = This.Blocks.Exists(Name)
End Function

Public Property Get BlocksArray() As Variant
  BlocksArray = This.Blocks.Items
End Property

Public Property Get ToDictionary() As Scripting.Dictionary
  Set BlocksDictionary = This.Blocks
End Property

Public Property Get Self() As CategorySheetBlocks
  Set Self = Me
End Property

'===============================================================================

Private Sub SortBySizes(ByVal ModelRecord As IRecord, ByVal Shape As Shape)
  Dim HaveSubs As Boolean
  HaveSubs = AddModelToSuitableBlockIfCountNotEmpty _
             (ModelRecord!SubCount1, ModelRecord, Mobiles.SubColumn1, Shape)
  HaveSubs = HaveSubs Or AddModelToSuitableBlockIfCountNotEmpty _
                         (ModelRecord!SubCount2, ModelRecord, Mobiles.SubColumn2, Shape)
  HaveSubs = HaveSubs Or AddModelToSuitableBlockIfCountNotEmpty _
                         (ModelRecord!SubCount3, ModelRecord, Mobiles.SubColumn3, Shape)
  HaveSubs = HaveSubs Or AddModelToSuitableBlockIfCountNotEmpty _
                         (ModelRecord!SubCount4, ModelRecord, Mobiles.SubColumn4, Shape)
  If Not HaveSubs Then AddModelToSuitableBlockIfCountNotEmpty _
                       ModelRecord!Count, ModelRecord, 0, Shape
End Sub

Private Function AddModelToSuitableBlockIfCountNotEmpty _
                 (ByVal Count As String, ByVal ModelRecord As IRecord, _
                  ByVal AdditionalBlocksID As Long, ByVal Shape As Shape) As Boolean
  
  If Helpers.IsFieldEmpty(Count) Then Exit Function
  
  Dim Model As New structModel
  Set Model.Shape = Shape.Duplicate
  Model.Count = This.SortingStrategy.Count(Count)
  
  Dim Block As CategorySheetBlock
  Set Block = GetOrCreateBlock _
    (This.SortingStrategy.Key(ModelRecord, AdditionalBlocksID, Count, Shape), _
     ModelRecord, AdditionalBlocksID, Count)
  Block.Add Model
  
  AddModelToSuitableBlockIfCountNotEmpty = True
  
End Function

Private Function GetOrCreateBlock(ByVal Key As String, _
                                  ByVal ModelRecord As IRecord, _
                                  ByVal AdditionalBlocksID As Long, _
                                  ByVal Count As String _
                                  ) As CategorySheetBlock
  With This
    
    If This.Blocks.Exists(Key) Then
      Set GetOrCreateBlock = This.Blocks(Key)
      Exit Function
    End If
    
    Set GetOrCreateBlock = CategorySheetBlock.Create
    This.Blocks.Add Key, GetOrCreateBlock
        
    GetOrCreateBlock.Caption.Line1 = _
      .SortingStrategy.Caption(ModelRecord, AdditionalBlocksID, Count)
    GetOrCreateBlock.MaxPlacesInWidth = _
      .Category!MaxPlacesInWidth
  
  End With
End Function
