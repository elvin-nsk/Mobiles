VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IRecordSet

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
                       ) As IRecordSet
  With New RecordSet
    Set Create = .Init(FieldNames, KeyFieldName, RecordFactory)
  End With
End Function

Friend Function Init(ByVal FieldNames As Collection, _
                     ByVal KeyFieldName As String, _
                     ByVal RecordFactory As IRecordFactory _
                     ) As IRecordSet
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

Private Function IRecordSet_BuildRecord() As IRecordBuilder
  Set IRecordSet_BuildRecord = This.RecordFactory.Builder(This.FieldNames)
  Set PendingRecord = IRecordSet_BuildRecord.Events
End Function

Private Property Get IRecordSet_Count() As Long
  IRecordSet_Count = This.Records.Count
End Property

Private Function IRecordSet_FieldNameExists(ByVal Name As String) As Boolean
  Dim Item As Variant
  For Each Item In This.FieldNames
    If Item = Name Then
      IRecordSet_FieldNameExists = True
      Exit Function
    End If
  Next Item
End Function

Private Function IRecordSet_FilterLike(ByVal Pattern As String) As Collection
  Set IRecordSet_FilterLike = New Collection
  Dim Record As IRecord
  For Each Record In This.Records
    If Record.ContainsLike(Pattern) Then IRecordSet_FilterLike.Add Record
  Next Record
End Function

Private Function IRecordSet_FilterMatch(ByVal Value As Variant) As Collection
  Set IRecordSet_FilterMatch = New Collection
  Dim Record As IRecord
  For Each Record In This.Records
    If Record.Contains(Value) Then IRecordSet_FilterMatch.Add Record
  Next Record
End Function

Private Function IRecordSet_GetFields(ByVal Name As String) As Collection
  Set IRecordSet_GetFields = New Collection
  Dim Record As IRecord
  For Each Record In This.Records
    IRecordSet_GetFields.Add Record.Field(Name)
  Next Record
End Function

Private Function IRecordSet_KeyFieldExists(ByVal Value As Variant) As Boolean
  If Not IsKeyFieldValid(Value) Then ThrowKeyFieldNotPrimitive
  Dim Record As IRecord
  For Each Record In This.Records
    If Record.Field(This.KeyFieldName) = Value Then
      IRecordSet_KeyFieldExists = True
      Exit Function
    End If
  Next Record
End Function

Private Property Get IRecordSet_KeyFieldSet() As Boolean
  IRecordSet_KeyFieldSet = This.IsKeyFieldSet
End Property

Private Property Get IRecordSet_NewEnum() As EnumHelper
  With New EnumHelper
    Set .EnumVariant = This.Records.[_NewEnum]
    Set IRecordSet_NewEnum = .Self
  End With
End Property

Private Property Get IRecordSet_Record(ByVal NumberOrKeyField As Variant) As IRecord
  Dim Index As Long
  Index = FindRecordIndex(NumberOrKeyField)
  If Index > 0 Then
    Set IRecordSet_Record = This.Records(Index)
  Else
    ThrowInvalidRecordNumberOrName
  End If
End Property

Private Function IRecordSet_RecordExists _
                 (ByVal NumberOrKeyField As Variant) As Boolean
  IRecordSet_RecordExists = Not (FindRecordIndex(NumberOrKeyField) = 0)
End Function

Private Property Get IRecordSet_Self() As IRecordSet
  Set IRecordSet_Self = Me
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
  VBA.Err.Raise CustomError, "RecordSet", "DuplicateFieldName"
End Sub

Private Sub ThrowKeyFieldsAreNotSet()
  VBA.Err.Raise CustomError, "RecordSet", "KeyFieldsAreNotSet"
End Sub

Private Sub ThrowInvalidRecordNumberOrName()
  VBA.Err.Raise CustomError, "RecordSet", "InvalidRecordNumberOrName"
End Sub

Private Sub ThrowKeyFieldNotPrimitive()
  VBA.Err.Raise CustomError, "RecordSet", "KeyFieldNotPrimitive"
End Sub

Private Sub ThrowKeyFieldDuplicate()
  VBA.Err.Raise CustomError, "RecordSet", "KeyFieldDuplicate"
End Sub

Private Property Get Self() As IRecordSet
  Set Self = Me
End Property
