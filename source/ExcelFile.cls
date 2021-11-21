VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Implements ITableFile

#Const UseLateBinding = False

'===============================================================================

#If UseLateBinding Then

  Private Type typeThis
    File As IFileSpec
    SheetIndex As Long
    StartingRow As Long
    StartingColumn As Long
    RowOffset As Long
    ColumnOffset As Long
    WorkbookIsNew As Boolean
    ExcelApp As Object
    Workbook As Object
    Table As Object
  End Type
  
#Else

  Private Type typeThis
    File As IFileSpec
    SheetIndex As Long
    StartingRow As Long
    StartingColumn As Long
    RowOffset As Long
    ColumnOffset As Long
    WorkbookIsNew As Boolean
    ExcelApp As Excel.Application
    Workbook As Excel.Workbook
    Table As Excel.Worksheet
  End Type

#End If

Private This As typeThis

'===============================================================================

Public Function Create(ByVal FileToBind As IFileSpec, _
                       Optional ByVal SheetIndex As Long = 1, _
                       Optional ByVal StartingRow As Long = 1, _
                       Optional ByVal StartingColumn As Long = 1) As ITableFile
  With New ExcelFile
    Set Create = .Init(FileToBind, SheetIndex, StartingRow, StartingColumn)
  End With
End Function

Private Sub Class_Terminate()
  With This
    If Not .File Is Nothing Then
      Save
      .Workbook.Close
    End If
  End With
End Sub

Friend Function Init(ByVal FileToBind As IFileSpec, _
                     Optional ByVal SheetIndex As Long = 1, _
                     Optional ByVal StartingRow As Long = 1, _
                     Optional ByVal StartingColumn As Long = 1) As ITableFile
  With This
    Set .File = FileToBind
    .SheetIndex = SheetIndex
    .StartingRow = StartingRow
    .StartingColumn = StartingColumn
    .RowOffset = .StartingRow - 1
    .ColumnOffset = .StartingColumn - 1
    #If UseLateBinding Then
      Set .ExcelApp = VBA.CreateObject("Excel.Application")
    #Else
      Set .ExcelApp = Excel.Application
    #End If
  End With
  Load
  Set Init = Me
End Function

'===============================================================================

Private Property Let ITableFile_Cell(ByVal Row As Long, _
                                     ByVal Column As Long, _
                                     RHS As String)
  With This
    .Table.Cells(Row + .RowOffset, Column + .ColumnOffset) = RHS
  End With
End Property
Private Property Get ITableFile_Cell(ByVal Row As Long, _
                                     ByVal Column As Long _
                                     ) As String
  With This
    ITableFile_Cell = .Table.Cells _
      (Row + .RowOffset, Column + .ColumnOffset)
  End With
End Property

Private Sub ITableFile_ForceClose()
  This.Workbook.Close
End Sub

Private Sub ITableFile_ForceLoad()
  Load
End Sub

Private Sub ITableFile_ForceSave()
  Save
End Sub

Private Property Get ITableFile_Self() As ITableFile
  Set ITableFile_Self = Me
End Property

'===============================================================================

Private Sub Load()
  With This
    If FileExists(.File.ToString) Then
      Set .Workbook = .ExcelApp.Workbooks.Open(.File.ToString)
    Else
      Set .Workbook = .ExcelApp.Workbooks.Add
      .WorkbookIsNew = True
    End If
    Set .Table = .Workbook.Sheets.Item(.SheetIndex)
  End With
End Sub

Private Sub Save()
  With This
    If .WorkbookIsNew Then
      .Workbook.SaveAs .File.ToString
    Else
      If Not .Workbook.Saved Then .Workbook.Save
    End If
  End With
End Sub

Public Function FileExists(ByVal File As String) As Boolean
  If File = "" Then Exit Function
  FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Private Property Get Self() As ITableFile
  Set Self = Me
End Property
