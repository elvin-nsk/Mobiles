VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupingStrategyBySize"
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
  With New GroupingStrategyBySize
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
                      ByVal Count As String) As String
  With This
    If AdditionalBlocksID = 0 Then
      If .Sizes.PrimaryFieldExists(Model!Size) Then
        IGroupingStrategy_Caption = .Sizes(Model!Size)("Name")
      Else
        IGroupingStrategy_Caption = Model!Size
      End If
    Else
      If .AdditionalBlocks.PrimaryFieldExists(VBA.CStr(AdditionalBlocksID)) Then
        IGroupingStrategy_Caption = _
          .AdditionalBlocks(VBA.CStr(AdditionalBlocksID))("Name")
      Else
        IGroupingStrategy_Caption = Model!Size
      End If
    End If
  End With
End Property

Private Property Get IGroupingStrategy_Count(ByVal Count As String) As Long
  IGroupingStrategy_Count = VBA.CLng(Count)
End Property

Private Property Get IGroupingStrategy_Key _
                     (ByVal Model As IRecord, _
                      ByVal AdditionalBlocksID As Long, _
                      ByVal Count As String, _
                      ByVal Shape As Shape _
                      ) As String
  If Model!Size = "" Then
    IGroupingStrategy_Key = GetSizeFromShape(Shape)
  Else
    If AdditionalBlocksID = 0 Then
      IGroupingStrategy_Key = Model!Size
    Else
      IGroupingStrategy_Key = Model!Size & "_" & VBA.CStr(AdditionalBlocksID)
    End If
  End If
End Property

'===============================================================================

Private Function GetSizeFromShape(ByVal Shape As Shape) As String
  GetSizeFromShape = VBA.CStr(VBA.Round(Shape.SizeWidth, 0))
  GetSizeFromShape = GetSizeFromShape & Mobiles.SizesMultiplierSymbol & _
                  VBA.CStr(VBA.Round(Shape.SizeHeight, 0))
End Function
