VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "structMobile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public Collation As structCollation
Public File As IFileSpec
Public Count As Long

'===============================================================================

Public Property Get Self() As structMobile
  Set Self = Me
End Property
