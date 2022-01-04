VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IRecordList

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Private Type typeThis
  Records As Collection
  FieldNames As Collection
  KeyFieldName As String
  IsKeyFieldSet As Boolean
  RecordFactory As IRecordFactory
End Type
Private This As typeThis
Private WithEvents PendingRecord As IRecordBuilderEvents
Attribute PendingRecord.VB_VarHelpID = -1

'===============================================================================

Public Function Create(ByVal FieldNames As Collection, _
                       Optional ByVal KeyFieldName As String, _
                       Optional ByVal RecordFactory As IRecordFactory _
                       ) As IRecordList
  With New RecordList
    Set Create = .Init(FieldNames, KeyFieldName, RecordFactory)
  End With
End Function

Friend Function Init(ByVal FieldNames As Collection, _
                     ByVal KeyFieldName As String, _
                     ByVal RecordFactory As IRecordFactory _
                     ) As IRecordList
  With This
    If RecordFactory Is Nothing Then
      Set .RecordFactory = Record
    Else
      Set .RecordFactory = RecordFactory
    End If
    SetFieldsAndCheck FieldNames
    SetKeyFieldName KeyFieldName
    Set .Records = New Collection
  End With
  Set Init = Me
End Function

'===============================================================================

Private Function IRecordList_BuildRecord() As IRecordBuilder
  Set IRecordList_BuildRecord = This.RecordFactory.Builder(This.FieldNames)
  Set PendingRecord = IRecordList_BuildRecord.Events
End Function

Private Property Get IRecordList_Count() As Long
  IRecordList_Count = This.Records.Count
End Property

Private Function IRecordList_FieldNameExists(ByVal Name As String) As Boolean
  Dim Item As Variant
  For Each Item In This.FieldNames
    If Item = Name Then
      IRecordList_FieldNameExists = True
      Exit Function
    End If
  Next Item
End Function

Private Function IRecordList_FilterLike(ByVal Pattern As String) As Collection
  Set IRecordList_FilterLike = New Collection
  Dim Record As IRecord
  For Each Record In This.Records
    If Record.ContainsLike(Pattern) Then IRecordList_FilterLike.Add Record
  Next Record
End Function

Private Function IRecordList_FilterMatch(ByVal Value As Variant) As Collection
  Set IRecordList_FilterMatch = New Collection
  Dim Record As IRecord
  For Each Record In This.Records
    If Record.Contains(Value) Then IRecordList_FilterMatch.Add Record
  Next Record
End Function

Private Function IRecordList_GetFields(ByVal Name As String) As Collection
  Set IRecordList_GetFields = New Collection
  Dim Record As IRecord
  For Each Record In This.Records
    IRecordList_GetFields.Add Record.Field(Name)
  Next Record
End Function

Private Function IRecordList_KeyFieldExists(ByVal Value As Variant) As Boolean
  If Not IsKeyFieldValid(Value) Then ThrowKeyFieldNotPrimitive
  Dim Record As IRecord
  For Each Record In This.Records
    If Record.Field(This.KeyFieldName) = Value Then
      IRecordList_KeyFieldExists = True
      Exit Function
    End If
  Next Record
End Function

Private Property Get IRecordList_KeyFieldSet() As Boolean
  IRecordList_KeyFieldSet = This.IsKeyFieldSet
End Property

Private Property Get IRecordList_NewEnum() As EnumHelper
  With New EnumHelper
    Set .EnumVariant = This.Records.[_NewEnum]
    Set IRecordList_NewEnum = .Self
  End With
End Property

Private Property Get IRecordList_Record(ByVal NumberOrKeyField As Variant) As IRecord
  Dim Index As Long
  Index = FindRecordIndex(NumberOrKeyField)
  If Index > 0 Then
    Set IRecordList_Record = This.Records(Index)
  Else
    ThrowInvalidRecordNumberOrName
  End If
End Property

Private Function IRecordList_RecordExists _
                 (ByVal NumberOrKeyField As Variant) As Boolean
  IRecordList_RecordExists = Not (FindRecordIndex(NumberOrKeyField) = 0)
End Function

Private Property Get IRecordList_Self() As IRecordList
  Set IRecordList_Self = Me
End Property

'===============================================================================

Private Sub PendingRecord_Built(ByVal Record As IRecord)
  CheckRecordKeyFieldOrThrow Record
  This.Records.Add Record
  Set PendingRecord = Nothing
End Sub

'===============================================================================

Private Sub SetKeyFieldName(ByVal KeyFieldName As String)
  With This
    If KeyFieldName = "" Then Exit Sub
    .IsKeyFieldSet = True
    .KeyFieldName = KeyFieldName
  End With
End Sub

Private Function IsKeyFieldValid(ByVal Value As Variant) As Boolean
  IsKeyFieldValid = VBA.IsNumeric(Value) Or _
                    VBA.VarType(Value) = vbString
End Function

Private Sub CheckRecordKeyFieldOrThrow(ByVal Record As IRecord)
  With This
    If Not .IsKeyFieldSet Then Exit Sub
    If Not IsKeyFieldValid(Record.Field(.KeyFieldName)) Then _
      ThrowKeyFieldNotPrimitive
    If Self.KeyFieldExists(Record.Field(.KeyFieldName)) Then _
      ThrowKeyFieldDuplicate
  End With
End Sub

Private Sub SetFieldsAndCheck(ByVal FieldNames As Collection)
  Set This.FieldNames = New Collection
  Dim Item As Variant
  For Each Item In FieldNames
    If Not Self.FieldNameExists(Item) Then
      This.FieldNames.Add Item
    Else
      ThrowDuplicateFieldName
    End If
  Next Item
End Sub

Private Function FindRecordIndex(ByVal NumberOrKeyField As Variant) As Long
  If This.Records.Count = 0 Then Exit Function
  If VBA.VarType(NumberOrKeyField) = vbString Then
    If Not This.IsKeyFieldSet Then ThrowKeyFieldsAreNotSet
    Dim Index As Long
    For Index = 1 To This.Records.Count
      If This.Records(Index).Field(This.KeyFieldName) = NumberOrKeyField Then
        FindRecordIndex = Index
        Exit Function
      End If
    Next Index
  ElseIf VBA.IsNumeric(NumberOrKeyField) Then
    If This.Records.Count >= VBA.CLng(NumberOrKeyField) Then
      FindRecordIndex = VBA.CLng(NumberOrKeyField)
    End If
  End If
End Function

Private Sub ThrowDuplicateFieldName()
  VBA.Err.Raise CustomError, "RecordList", "DuplicateFieldName"
End Sub

Private Sub ThrowKeyFieldsAreNotSet()
  VBA.Err.Raise CustomError, "RecordList", "KeyFieldsAreNotSet"
End Sub

Private Sub ThrowInvalidRecordNumberOrName()
  VBA.Err.Raise CustomError, "RecordList", "InvalidRecordNumberOrName"
End Sub

Private Sub ThrowKeyFieldNotPrimitive()
  VBA.Err.Raise CustomError, "RecordList", "KeyFieldNotPrimitive"
End Sub

Private Sub ThrowKeyFieldDuplicate()
  VBA.Err.Raise CustomError, "RecordList", "KeyFieldDuplicate"
End Sub

Private Property Get Self() As IRecordList
  Set Self = Me
End Property
