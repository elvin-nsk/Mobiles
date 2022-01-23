VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGroupingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
Option Explicit

'===============================================================================

Public Property Get Caption(ByVal Model As IRecord, _
                            ByVal AdditionalBlocksID As Long, _
                            ByVal Count As String) As String
End Property

Public Property Get Count(ByVal Count As String) As Long
End Property

Public Property Get Key(ByVal Model As IRecord, _
                        ByVal AdditionalBlocksID As Long, _
                        ByVal Count As String, _
                        ByVal Shape As Shape _
                        ) As String
End Property
