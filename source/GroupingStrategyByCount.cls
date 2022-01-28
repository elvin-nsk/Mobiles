VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupingStrategyByCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Implements IGroupingStrategy

'===============================================================================

Private Type typeThis
  Category As IRecord
  AdditionalBlocks As IRecordList
  Sizes As IRecordList
End Type
Private This As typeThis

'===============================================================================

Public Function Create( _
                  ByVal Category As IRecord, _
                  ByVal AdditionalBlocks As IRecordList, _
                  ByVal Sizes As IRecordList _
                ) As IGroupingStrategy
  With New GroupingStrategyByCount
    Set Create = .Init(Category, AdditionalBlocks, Sizes)
  End With
End Function

Friend Function Init( _
                  ByVal Category As IRecord, _
                  ByVal AdditionalBlocks As IRecordList, _
                  ByVal Sizes As IRecordList _
                ) As IGroupingStrategy
  With This
    Set .Category = Category
    Set .AdditionalBlocks = AdditionalBlocks
    Set .Sizes = Sizes
  End With
  Set Init = Me
End Function

'===============================================================================

Private Property Get IGroupingStrategy_Caption _
                     (ByVal Model As IRecord, _
                      ByVal AdditionalBlocksID As Long, _
                      ByVal Count As String, _
                      ByVal Shape As Shape _
                      ) As String
  IGroupingStrategy_Caption = GetCountText(Count)
End Property

Private Property Get IGroupingStrategy_Count(ByVal CountString As String) As Long
  IGroupingStrategy_Count = 1
End Property

Private Property Get IGroupingStrategy_Key _
                     (ByVal Model As IRecord, _
                      ByVal AdditionalBlocksID As Long, _
                      ByVal Count As String, _
                      ByVal Shape As Shape _
                      ) As String
  IGroupingStrategy_Key = Self.Caption(Model, AdditionalBlocksID, Count, Shape)
End Property

'===============================================================================

Private Function GetCountText(ByVal Count As String) As String
  GetCountText = "ןמ " & Count & " רע."
End Function

Private Property Get Self() As IGroupingStrategy
  Set Self = Me
End Property
