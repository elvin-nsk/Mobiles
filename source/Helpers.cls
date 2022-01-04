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

Public Function GetExcelFile() As IEither
  Dim File As IFileSpec
  Dim Path As String
  Path = GetSetting _
    (AppName, Section, Key, "%USERPROFILE%\Desktop")
  Dim FilesCollection As New Collection
  With FileBrowser.Create
    .Filter = "Excel Files (*.xlsx)" & VBA.Chr(0) & "*.xlsx"
    .InitialDir = Path
    .MultiSelect = False
    .Title = "Выберите таблицу категорий"
    Set FilesCollection = .ShowFileOpenDialog
  End With
  If FilesCollection.Count = 0 Then
    Set GetExcelFile = Either.Create()
    Exit Function
  End If
  Set File = FileSpec.Create(FilesCollection(1))
  Set GetExcelFile = Either.Create(File)
  Path = File.Path
  SaveSetting AppName, Section, Key, Path
End Function

Public Function OpenTableFile(ByVal File As IFileSpec, _
                              Optional ByVal ReadOnly As Boolean = False _
                              ) As IEither
  Dim Result As ITableFile
  On Error GoTo ErrorHandler
  If ReadOnly Then
    Set Result = ExcelConnection.CreateReadOnly(FileToBind:=File, SheetName:="Лист1", StartingRow:=2)
  Else
    Set Result = ExcelConnection.Create(FileToBind:=File, SheetName:="Лист1", StartingRow:=2)
  End If
  Set OpenTableFile = Either.Create(Result)
  Exit Function
ErrorHandler:
  Set OpenTableFile = Either.Create()
End Function

Public Function BindMainTable(ByVal Table As ITableFile, _
                              ByVal KeyFieldName As String _
                              ) As IRecordListToTableBinder
  With RecordListToTableBinder.Builder(Table)
    .WithField "Name", 1
    .WithField "Count", 2
    .WithField "File", 3
    .WithKeyField KeyFieldName
    Set BindMainTable = .Build
  End With
End Function

Public Sub ResetMobilesCount(ByVal RecordList As IRecordList)
  Dim Record As IRecord
  For Each Record In RecordList.NewEnum
    Record.Field("Count") = 0
  Next Record
End Sub

Public Sub CountMobilesInShapes(ByVal RecordSet As IRecordList, _
                                ByVal Shapes As ShapeRange)
  Dim PBar As IProgressBar
  Set PBar = ProgressBar.CreateNumeric(Shapes.Count)
  PBar.Caption = "Подсчёт мобайлов..."
  Dim Name As String
  Dim Shape As Shape
  For Each Shape In Shapes
    Name = Shape.Name
    If Not Name = "" Then
      If RecordSet.KeyFieldExists(Name) Then
        PBar.Update
        RecordSet.Record(Name)("Count") = RecordSet.Record(Name)("Count") + 1
      End If
    End If
  Next Shape
End Sub

Public Sub DebugPathsReplace(ByVal MobilesDic As Dictionary)
  Dim Mobile As structMobile
  Dim Key As Variant
  For Each Key In MobilesDic.Keys
    MobilesDic(Key).File.Path = _
      VBA.Replace(MobilesDic(Key).File.Path, _
        Mobiles.DebugMobilesRootRepalceFrom, _
        Mobiles.DebugMobilesRootRepalceTo)
  Next Key
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
