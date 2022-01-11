VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Helpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Singleton
'@PredeclaredId
Option Explicit

'===============================================================================

'константы для сохранения последнего пути в реестре
Const AppName As String = "Mobiles"
Const Section As String = "Settings"
Const Key As String = "TableLastPath"

Private Sub Class_Initialize()
  If Not Me Is Helpers Then Err.Raise 425 'Invalid object use
End Sub

'===============================================================================

Public Function tryGetExcelFile() As IEither
  Dim File As IFileSpec
  Dim Path As String
  Path = GetSetting _
         (AppName, Section, Key, "%USERPROFILE%\Desktop")
  Dim FilesCollection As New Collection
  With FileBrowser.Create
    .Filter = "Excel Files (*.xlsx, .xls)" & VBA.Chr(0) & "*.xlsx;*.xls"
    .InitialDir = Path
    .MultiSelect = False
    .Title = "Выберите таблицу"
    Set FilesCollection = .ShowFileOpenDialog
  End With
  If FilesCollection.Count = 0 Then
    Set tryGetExcelFile = Either.Create()
    Exit Function
  End If
  Set File = FileSpec.Create(FilesCollection(1))
  Set tryGetExcelFile = Either.Create(File)
  Path = File.Path
  SaveSetting AppName, Section, Key, Path
End Function

Public Function tryBindMainTable(ByVal File As IFileSpec, _
                                 ByVal NameIsPrimaryKey As Boolean, _
                                 ByVal ReadOnly As Boolean _
                                 ) As IEither
  Dim Table As ITableFile
  Set Table = OpenTableFileOrNotify _
      (File:=File, SheetName:=Mobiles.MainTableName, ReadOnly:=ReadOnly)
  If Table Is Nothing Then
    Set tryBindMainTable = Either.Create
    Exit Function
  End If
  
  With RecordListToTableBinder.Builder(Table)
    .WithKey "Size", 1
    If NameIsPrimaryKey Then .WithPrimaryKey "Name", 2 Else .WithKey "Name", 2
    .WithKey "Count", 3
    .WithMandatoryKey "File", 4
    .WithKey "SubCount1", 5
    .WithKey "SubCount2", 6
    .WithKey "SubCount3", 7
    .WithKey "SubCount4", 8
    '.WithUnboundKey "Valid"
    Set tryBindMainTable = Either.Create(.Build)
  End With
End Function

Public Function tryBindCategoriesTable(ByVal File As IFileSpec) As IEither
  Dim Table As ITableFile
  Set Table = OpenTableFileOrNotify _
      (File:=File, SheetName:=Mobiles.CategoriesTableName, ReadOnly:=True)
  If Table Is Nothing Then
    Set tryBindCategoriesTable = Either.Create
    Exit Function
  End If
  With RecordListToTableBinder.Builder(Table)
    .WithPrimaryKey "Name", 1
    .WithKey "Sizes", 2
    .WithKey "Line1", 3
    .WithKey "Line2", 4
    .WithKey "MaxBlocksInWidth", 5
    .WithKey "MaxBlocksInHeight", 6
    .WithKey "MaxPlacesInWidth", 7
    .WithKey "MaxPlacesInHeight", 8
    Set tryBindCategoriesTable = Either.Create(.Build)
  End With
End Function

Public Function tryBindSubTable(ByVal File As IFileSpec) As IEither
  Dim Table As ITableFile
  Set Table = OpenTableFileOrNotify _
      (File:=File, SheetName:=Mobiles.SubTableName, ReadOnly:=True)
  If Table Is Nothing Then
    Set tryBindSubTable = Either.Create
    Exit Function
  End If
  With RecordListToTableBinder.Builder(Table)
    .WithPrimaryKey "Column", 1
    .WithKey "Name", 2
    Set tryBindSubTable = Either.Create(.Build)
  End With
End Function

Public Function tryBindSizesTable(ByVal File As IFileSpec) As IEither
  Dim Table As ITableFile
  Set Table = OpenTableFileOrNotify _
      (File:=File, SheetName:=Mobiles.SizesTableName, ReadOnly:=True)
  If Table Is Nothing Then
    Set tryBindSizesTable = Either.Create
    Exit Function
  End If
  With RecordListToTableBinder.Builder(Table)
    .WithPrimaryKey "Size", 1
    .WithKey "Name", 2
    Set tryBindSizesTable = Either.Create(.Build)
  End With
End Function

Public Function OpenTableFileOrNotify(ByVal File As IFileSpec, _
                                      ByVal SheetName As String, _
                                      ByVal ReadOnly As Boolean _
                                      ) As ITableFile
  On Error GoTo ErrorHandler
  If ReadOnly Then
    Set OpenTableFileOrNotify = ExcelConnection.CreateReadOnly _
        (FileToBind:=File, SheetName:=SheetName, StartingRow:=2)
  Else
    Set OpenTableFileOrNotify = ExcelConnection.Create _
        (FileToBind:=File, SheetName:=SheetName, StartingRow:=2)
  End If
  Exit Function
ErrorHandler:
  VBA.MsgBox "Ошибка чтения файла", vbCritical
  On Error GoTo 0
End Function

Public Sub ResetMobilesCount(ByVal RecordList As IRecordList)
  Dim Record As IRecord
  For Each Record In RecordList.NewEnum
    Record.Field("Count") = 0
  Next Record
End Sub

Public Sub ValidateMainTable(ByVal RecordList As IRecordList)
  Dim Record As IRecord
  For Each Record In RecordList.NewEnum
    ValidateMainTableRecord Record
  Next Record
End Sub

Public Sub ValidateMainTableRecord(ByVal Record As IRecord)
  If Record.Field("Size") = "" Then Exit Sub
  If Record.Field("File") = "" Then Exit Sub
  Record.Field("Valid") = True
End Sub

Public Sub CountMobilesInShapes(ByVal RecordList As IRecordList, _
                                ByVal Shapes As ShapeRange)
  Dim PBar As IProgressBar
  Set PBar = ProgressBar.CreateNumeric(Shapes.Count)
  PBar.Caption = "Подсчёт мобайлов..."
  Dim Name As String
  Dim Shape As Shape
  For Each Shape In Shapes
    Name = Shape.Name
    If Not Name = "" Then
      If RecordList.RecordExists(Name) Then
        PBar.Update
        RecordList.Record(Name)("Count") = RecordList.Record(Name)("Count") + 1
      End If
    End If
  Next Shape
End Sub

Public Sub DebugPathsReplace(ByVal Records As IRecordList)
  Const RecordName As String = "File"
  Dim Record As IRecord
  For Each Record In Records.NewEnum
    Record(RecordName) = VBA.Replace _
      (Record(RecordName), _
       Mobiles.DebugMobilesRootRepalceFrom, _
       Mobiles.DebugMobilesRootRepalceTo)
  Next Record
End Sub

Public Sub Report(ByVal FailedFiles As Collection)
  Dim Message As String
  Message = "Следующие файлы не найдены:" & vbCrLf
  Dim i As Long
  For i = 1 To FailedFiles.Count
    Message = Message & FailedFiles(i) & vbCrLf
  Next i
  VBA.MsgBox Message, vbInformation, "Ошибка импорта"
End Sub
