VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordSetToTableBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IRecordSetToTableBinder
Implements IRecordSetToTableBinderBuilder

'===============================================================================

Private Enum enumAction
  ActionCreateAndRead
  ActionRead
  ActionWrite
End Enum

Private Type typeThis
  RecordSet As IRecordSet
  Table As ITableFile
  Bindings As Scripting.Dictionary
  RecordsToRowsCollation As Collection
  KeyFieldName As String
  IsKeyFieldSet As Boolean
  FieldNames As Collection
  RecordFactory As IRecordFactory
End Type
Private This As typeThis

'===============================================================================

Public Function Builder(ByVal Table As ITableFile, _
                        Optional ByVal RecordFactory As IRecordFactory _
                        ) As IRecordSetToTableBinderBuilder
  With New RecordSetToTableBinder
    Set Builder = .Init(Table, RecordFactory)
  End With
End Function

Private Sub Class_Terminate()
  If Not This.RecordSet Is Nothing And _
     Not This.Table Is Nothing Then _
    Self.ForceFlush
End Sub

Friend Function Init(ByVal Table As ITableFile, _
                     ByVal RecordFactory As IRecordFactory _
                     ) As IRecordSetToTableBinderBuilder
  With This
    Set .Table = Table
    Set .Bindings = New Scripting.Dictionary
    Set .RecordsToRowsCollation = New Collection
    Set .RecordFactory = RecordFactory
    Set .FieldNames = New Collection
  End With
  Set Init = Me
End Function

'-------------------------------------------------------------------------------

Private Function IRecordSetToTableBinderBuilder_WithField _
                 (ByVal FieldName As String, ByVal ColumnToBind As Long) As IRecordSetToTableBinderBuilder
  This.Bindings.Add FieldName, ColumnToBind
  This.FieldNames.Add FieldName
End Function

Private Function IRecordSetToTableBinderBuilder_WithKeyField _
                 (ByVal KeyFieldName As String) As IRecordSetToTableBinderBuilder
  This.KeyFieldName = KeyFieldName
End Function

Private Function IRecordSetToTableBinderBuilder_Build() As IRecordSetToTableBinder
  With This
    .IsKeyFieldSet = Not (.KeyFieldName = "")
    Set .RecordSet = RecordSet.Create(.FieldNames, .KeyFieldName, .RecordFactory)
    TableWalk ActionCreateAndRead
    Set IRecordSetToTableBinderBuilder_Build = Me
    Set .FieldNames = Nothing
    Set .RecordFactory = Nothing
  End With
End Function

'===============================================================================

Private Property Get IRecordSetToTableBinder_RecordSet() As IRecordSet
  Set IRecordSetToTableBinder_RecordSet = This.RecordSet
End Property

Private Sub IRecordSetToTableBinder_FlushAndCloseTable()
  Self.ForceFlush
  This.Table.ForceClose
End Sub

Private Sub IRecordSetToTableBinder_ForceFlush()
  If Not This.Table.ReadOnly Then
    TableWalk ActionWrite
    This.Table.ForceSave
  End If
End Sub

Private Property Get IRecordSetToTableBinder_Self() As IRecordSetToTableBinder
  Set IRecordSetToTableBinder_Self = Me
End Property

'===============================================================================

Private Sub TableWalk(ByVal Action As enumAction)
  Dim Row As Long
  For Row = 1 To This.Table.MaxRow
    ProcessRow Row, Action
  Next Row
End Sub

Private Sub ProcessRow(ByVal Row As Long, ByVal Action As enumAction)
  
  Dim RecordBuilder As IRecordBuilder
  Dim RecordIndex As Long
  
  If Action = ActionCreateAndRead Then
    If Not IsRowValidToRecord(Row) Then Exit Sub
    Set RecordBuilder = This.RecordSet.BuildRecord
    This.RecordsToRowsCollation.Add Row
  ElseIf Action = ActionRead Then
    RecordIndex = FindRecordIndex(Row)
    If RecordIndex = 0 Then Exit Sub
  ElseIf Action = ActionWrite Then
    RecordIndex = FindRecordIndex(Row)
    If RecordIndex = 0 Then Exit Sub
    If Not This.RecordSet(RecordIndex).IsChanged Then Exit Sub
  End If
  
  Dim Key As Variant
  For Each Key In This.Bindings
    
    If Action = ActionCreateAndRead Then
      RecordBuilder.WithField Key, This.Table.Cell(Row, This.Bindings(Key))
    ElseIf Action = ActionRead Then
      This.RecordSet(RecordIndex)(Key) = This.Table.Cell(Row, This.Bindings(Key))
    ElseIf Action = ActionWrite Then
      This.Table.Cell(Row, This.Bindings(Key)) = This.RecordSet(RecordIndex)(Key)
    End If
  
  Next Key
  
  If Action = ActionCreateAndRead Then RecordBuilder.Build
  
End Sub

Private Function IsRowValidToRecord(ByVal Row As Long) As Boolean
  With This
    If .IsKeyFieldSet Then
      If .Table.Cell(Row, .Bindings(.KeyFieldName)) = "" Then Exit Function
    End If
    IsRowValidToRecord = True
  End With
End Function

Private Function FindRecordIndex(ByVal Row As Long) As Long
  If This.RecordsToRowsCollation.Count = 0 Then Exit Function
  Dim i As Long
  For i = 1 To This.RecordsToRowsCollation.Count
    If This.RecordsToRowsCollation(i) = Row Then
      FindRecordIndex = i
      Exit Function
    End If
  Next i
End Function

Private Property Get Self() As IRecordSetToTableBinder
  Set Self = Me
End Property
