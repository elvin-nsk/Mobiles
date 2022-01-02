VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "structCollation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public ID As Long
Public Name As String
Public SheetLine1 As String
Public SheetLine2 As String

'===============================================================================

Public Property Get Self() As structCollation
  Set Self = Me
End Property
