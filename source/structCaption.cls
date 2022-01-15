VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "structCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public Line1 As String
Public Line2 As String
Public FontSize As Double
Public Color As String
Public Line1Bold As Boolean

'===============================================================================

Public Property Get Self() As structCaption
  Set Self = Me
End Property
