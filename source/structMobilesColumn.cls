VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "structMobilesColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public Count As Long
Public Shapes As ShapeRange
Public CaptionShape As Shape

'===============================================================================

Public Property Get Self() As structMobilesColumn
  Set Self = Me
End Property
