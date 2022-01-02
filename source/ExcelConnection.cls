VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements ITableFile

'===============================================================================

Private Const adLockReadOnly = 1
Private Const adLockOptimistic = 3
Private Const adUseClient = 3
Private Const adSchemaTables = 20

Private Type typeThis
  File As IFileSpec
  SheetName As String
  StartingRow As Long
  StartingColumn As Long
  RowOffset As Long
  ColumnOffset As Long
  ReadOnly As Boolean
  Connection As Object
  RecordSet As Object
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal FileToBind As IFileSpec, _
                       Optional ByVal SheetName As String, _
                       Optional ByVal StartingRow As Long = 1, _
                       Optional ByVal StartingColumn As Long = 1 _
                       ) As ITableFile
  With New ExcelConnection
    Set Create = .Init(FileToBind, SheetName, StartingRow, StartingColumn)
  End With
End Function

Public Function CreateReadOnly(ByVal FileToBind As IFileSpec, _
                               Optional ByVal SheetName As String, _
                               Optional ByVal StartingRow As Long = 1, _
                               Optional ByVal StartingColumn As Long = 1 _
                               ) As ITableFile
  With New ExcelConnection
    Set CreateReadOnly = .Init(FileToBind, SheetName, StartingRow, StartingColumn, True)
  End With
End Function

Private Sub Class_Terminate()
  With This
    If Not .RecordSet Is Nothing Then
      Self.ForceSave
      .RecordSet.Close
    End If
    If Not .Connection Is Nothing Then .Connection.Close
  End With
End Sub

Friend Function Init(ByVal FileToBind As IFileSpec, _
                     ByVal SheetName As String, _
                     ByVal StartingRow As Long, _
                     ByVal StartingColumn As Long, _
                     Optional ByVal ReadOnly As Boolean = False _
                     ) As ITableFile
  With This
    Set .File = FileToBind
    .SheetName = SheetName
    .StartingRow = StartingRow
    .StartingColumn = StartingColumn
    .RowOffset = .StartingRow - 2
    .ColumnOffset = .StartingColumn - 2
    .ReadOnly = ReadOnly
    Connect
  End With
  Set Init = Me
End Function

'===============================================================================

Private Property Let ITableFile_Cell(ByVal Row As Long, _
                                     ByVal Column As Long, _
                                     RHS As String)
  With This.RecordSet
    If Not MoveToRow(Row) Then Exit Property
    If .EOF Or .BOF Then Exit Property
    .Fields(Column + This.ColumnOffset) = RHS
    .Update
  End With
End Property
Private Property Get ITableFile_Cell(ByVal Row As Long, _
                                     ByVal Column As Long _
                                     ) As String
  Dim TempVar As Variant
  With This.RecordSet
    If Not MoveToRow(Row) Then Exit Property
    TempVar = .Fields(Column + This.ColumnOffset)
    If Not VBA.IsNull(TempVar) Then
      ITableFile_Cell = TempVar
    End If
  End With
End Property

Private Sub ITableFile_ForceClose()
  With This
    If Not .RecordSet Is Nothing Then
      If Not .ReadOnly Then .RecordSet.CancelUpdate
      .RecordSet.Close
    End If
    If Not .Connection Is Nothing Then .Connection.Close
  End With
End Sub

Private Sub ITableFile_ForceLoad()

End Sub

Private Sub ITableFile_ForceSave()
  With This
    If Not .ReadOnly Then .RecordSet.Update
  End With
End Sub

Private Property Get ITableFile_MaxColumn() As Long
  ITableFile_MaxColumn = This.RecordSet.Fields.Count
End Property

Private Property Get ITableFile_MaxRow() As Long
  ITableFile_MaxRow = This.RecordSet.RecordCount
End Property

Private Property Get ITableFile_ReadOnly() As Boolean

End Property

Private Property Get ITableFile_Self() As ITableFile
  Set ITableFile_Self = Me
End Property

'===============================================================================

Private Sub Connect()
  Set This.Connection = VBA.CreateObject("ADODB.Connection")
  With This.Connection
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Properties("Extended Properties").Value = "Excel 12.0;HDR=No"
    .Open This.File.ToString
    ValidateSheetName
  End With
  Set This.RecordSet = VBA.CreateObject("ADODB.RecordSet")
  With This.RecordSet
    .ActiveConnection = This.Connection
    If This.ReadOnly Then
      .LockType = adLockReadOnly
    Else
      .LockType = adLockOptimistic
    End If
    .CursorLocation = adUseClient
    .Source = "SELECT * FROM [" & This.SheetName & "]"
    .Open
  End With
End Sub

Private Function ValidateSheetName()
  If This.SheetName = "" Then
    With This.Connection.OpenSchema(adSchemaTables)
      This.SheetName = .Fields("table_name").Value
      .Close
    End With
  Else
    This.SheetName = This.SheetName & "$"
  End If
End Function

Private Function MoveToRow(ByVal Row As Long) As Boolean
  With This.RecordSet
    .MoveFirst
    .Move Row + This.RowOffset
    If Not .EOF Then MoveToRow = True
  End With
End Function

Private Property Get Self() As ITableFile
  Set Self = Me
End Property