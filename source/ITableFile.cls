VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
Option Explicit

'===============================================================================

Public Property Get Cell(ByVal Row As Long, _
                         ByVal Column As Long) _
                         As String
End Property
Public Property Let Cell(ByVal Row As Long, _
                         ByVal Column As Long, _
                         RHS As String)
End Property

Public Sub ForceClose()
End Sub

Public Sub ForceLoad()
End Sub

Public Sub ForceSave()
End Sub

'===============================================================================

Public Property Get Self() As ITableFile
End Property
