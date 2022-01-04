VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordListToTableBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IRecordListToTableBinder
Implements IRecordListToTableBinderBuilder

'===============================================================================

Private Enum enumAction
  ActionCreateAndRead
  ActionRead
  ActionWrite
End Enum

Private Type typeThis
  RecordList As IRecordList
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
                        ) As IRecordListToTableBinderBuilder
  With New RecordListToTableBinder
    Set Builder = .Init(Table, RecordFactory)
  End With
End Function

Private Sub Class_Terminate()
  If Not This.RecordList Is Nothing And _
     Not This.Table Is Nothing Then _
    Self.ForceFlush
End Sub

Friend Function Init(ByVal Table As ITableFile, _
                     ByVal RecordFactory As IRecordFactory _
                     ) As IRecordListToTableBinderBuilder
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

Private Function IRecordListToTableBinderBuilder_WithField _
                 (ByVal FieldName As String, ByVal ColumnToBind As Long) As IRecordListToTableBinderBuilder
  This.Bindings.Add FieldName, ColumnToBind
  This.FieldNames.Add FieldName
End Function

Private Function IRecordListToTableBinderBuilder_WithKeyField _
                 (ByVal KeyFieldName As String) As IRecordListToTableBinderBuilder
  This.KeyFieldName = KeyFieldName
End Function

Private Function IRecordListToTableBinderBuilder_Build() As IRecordListToTableBinder
  With This
    .IsKeyFieldSet = Not (.KeyFieldName = "")
    Set .RecordList = RecordList.Create(.FieldNames, .KeyFieldName, .RecordFactory)
    TableWalk ActionCreateAndRead
    Set IRecordListToTableBinderBuilder_Build = Me
    Set .FieldNames = Nothing
    Set .RecordFactory = Nothing
  End With
End Function

'===============================================================================

Private Property Get IRecordListToTableBinder_RecordList() As IRecordList
  Set IRecordListToTableBinder_RecordList = This.RecordList
End Property

Private Sub IRecordListToTableBinder_FlushAndCloseTable()
  Self.ForceFlush
  This.Table.ForceClose
End Sub

Private Sub IRecordListToTableBinder_ForceFlush()
  If Not This.Table.ReadOnly Then
    TableWalk ActionWrite
    This.Table.ForceSave
  End If
End Sub

Private Property Get IRecordListToTableBinder_Self() As IRecordListToTableBinder
  Set IRecordListToTableBinder_Self = Me
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
    Set RecordBuilder = This.RecordList.BuildRecord
    This.RecordsToRowsCollation.Add Row
  ElseIf Action = ActionRead Then
    RecordIndex = FindRecordIndex(Row)
    If RecordIndex = 0 Then Exit Sub
  ElseIf Action = ActionWrite Then
    RecordIndex = FindRecordIndex(Row)
    If RecordIndex = 0 Then Exit Sub
    If Not This.RecordList(RecordIndex).IsChanged Then Exit Sub
  End If
  
  Dim Key As Variant
  For Each Key In This.Bindings
    
    If Action = ActionCreateAndRead Then
      RecordBuilder.WithField Key, This.Table.Cell(Row, This.Bindings(Key))
    ElseIf Action = ActionRead Then
      This.RecordList(RecordIndex)(Key) = This.Table.Cell(Row, This.Bindings(Key))
    ElseIf Action = ActionWrite Then
      This.Table.Cell(Row, This.Bindings(Key)) = This.RecordList(RecordIndex)(Key)
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

Private Property Get Self() As IRecordListToTableBinder
  Set Self = Me
End Property
