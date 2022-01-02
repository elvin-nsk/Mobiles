VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IRecord
Implements IRecordFactory
Implements IRecordBuilder

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Private Type typeThis
  Encapsulated As Scripting.Dictionary
  Changed As Boolean
  BuilderEvents As IRecordBuilderEvents
  BuilderAssignedNames As Collection
End Type
Private This As typeThis

'===============================================================================

Private Function IRecordFactory_Builder _
                 (ByVal FieldNames As Collection) As IRecordBuilder
  With New Record
    Set IRecordFactory_Builder = .Init(FieldNames)
  End With
End Function

Friend Function Init(ByVal FieldNames As Collection) As IRecordBuilder
  With This
    Set .Encapsulated = New Scripting.Dictionary
    Dim Item As Variant
    For Each Item In FieldNames
      .Encapsulated.Add Item, Empty
    Next Item
    Set .BuilderEvents = New IRecordBuilderEvents
    Set .BuilderAssignedNames = New Collection
  End With
  Set Init = Me
End Function

'-------------------------------------------------------------------------------

Private Function IRecordBuilder_WithField _
                 (ByVal Name As String, ByVal Value As Variant) As IRecordBuilder
  CheckNameOrThrow Name
  CheckAlreadyAssigned Name
  AssignUnknownItem Name, Value
  This.BuilderAssignedNames.Add Name
  Set IRecordBuilder_WithField = Me
End Function

Private Function IRecordBuilder_Build() As IRecord
  Set IRecordBuilder_Build = Me
  This.BuilderEvents.Built Me
  Set This.BuilderEvents = Nothing
  Set This.BuilderAssignedNames = Nothing
End Function

Private Property Get IRecordBuilder_Events() As IRecordBuilderEvents
  Set IRecordBuilder_Events = This.BuilderEvents
End Property

'===============================================================================

Private Function IRecord_Contains(ByVal Value As Variant) As Boolean
  Dim Item As Variant
  For Each Item In This.Encapsulated.Items
    If Item = Value Then
      IRecord_Contains = True
      Exit Function
    End If
  Next Item
End Function

Private Function IRecord_ContainsLike(ByVal Pattern As String) As Boolean
  Dim Item As Variant
  For Each Item In This.Encapsulated.Items
    If Item Like Pattern Then
      IRecord_ContainsLike = True
      Exit Function
    End If
  Next Item
End Function

Private Function IRecord_Exists(ByVal FieldName As String) As Boolean
  IRecord_Exists = This.Encapsulated.Exists(FieldName)
End Function

Private Property Get IRecord_Field(ByVal Name As String) As Variant
  CheckNameOrThrow Name
  AssignUnknown IRecord_Field, This.Encapsulated(Name)
End Property
Private Property Let IRecord_Field(ByVal Name As String, RHS As Variant)
  CheckNameOrThrow Name
  This.Encapsulated(Name) = RHS
  This.Changed = True
End Property
Private Property Set IRecord_Field(ByVal Name As String, RHS As Variant)
  CheckNameOrThrow Name
  Set This.Encapsulated(Name) = RHS
  This.Changed = True
End Property

Private Property Get IRecord_FieldNames() As Collection
  Set IRecord_FieldNames = New Collection
  Dim Key As Variant
  For Each Key In This.Encapsulated.Keys
    IRecord_FieldNames.Add Key
  Next Key
End Property

Private Property Get IRecord_IsChanged() As Boolean
  IRecord_IsChanged = This.Changed
End Property

Private Property Get IRecord_Self() As IRecord
  Set IRecord_Self = Me
End Property

'===============================================================================

Private Sub CheckAlreadyAssigned(ByVal Name As String)
  Dim Item As Variant
  For Each Item In This.BuilderAssignedNames
    If Item = Name Then ThrowFieldAlreadySet Name
  Next Item
End Sub

Private Sub CheckNameOrThrow(ByVal Name As String)
  If Not This.Encapsulated.Exists(Name) Then ThrowWrongFieldName
End Sub

Public Sub AssignUnknown(ByRef Destination As Variant, ByRef Value As Variant)
  If VBA.IsObject(Value) Then
    Set Destination = Value
  Else
    Destination = Value
  End If
End Sub

Public Sub AssignUnknownItem(ByVal Key As Variant, _
                             ByRef Value As Variant)
  If VBA.IsObject(Value) Then
    Set This.Encapsulated(Key) = Value
  Else
    This.Encapsulated(Key) = Value
  End If
End Sub

Private Sub ThrowFieldAlreadySet(ByVal Name As String)
  VBA.Err.Raise CustomError, "Record", "Значение поля '" & Name & "' уже задано"
End Sub

Private Sub ThrowWrongFieldName()
  VBA.Err.Raise CustomError, "Record", "WrongFieldName"
End Sub

Private Property Get Self() As IRecord
  Set Self = Me
End Property
