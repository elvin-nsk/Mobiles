VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecordListToTableBinderBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Function WithField(ByVal FieldName As String, _
                          ByVal ColumnToBind As Long _
                          ) As IRecordListToTableBinderBuilder
End Function

Public Function WithKeyField(ByVal KeyFieldName As String) As IRecordListToTableBinderBuilder
End Function

'===============================================================================

Public Function Build() As IRecordListToTableBinder
End Function
