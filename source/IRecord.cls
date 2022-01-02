VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

'@DefaultMember
Public Property Get Field(ByVal Name As String) As Variant
Attribute Field.VB_UserMemId = 0
End Property
Public Property Let Field(ByVal Name As String, RHS As Variant)
End Property
Public Property Set Field(ByVal Name As String, RHS As Variant)
End Property

Public Property Get FieldNames() As Collection
End Property

Public Function Exists(ByVal FieldName As String) As Boolean
End Function

Public Function Contains(ByVal Value As Variant) As Boolean
End Function

Public Function ContainsLike(ByVal Pattern As String) As Boolean
End Function

Public Property Get IsChanged() As Boolean
End Property

'===============================================================================

Public Property Get Self() As IRecord
End Property