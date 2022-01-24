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

Public Function tryBindModelsTable(ByVal File As IFileSpec, _
                                   ByVal NameIsPrimaryKey As Boolean, _
                                   ByVal ReadOnly As Boolean _
                                   ) As IEither
  Dim Table As ITableFile
  Set Table = OpenTableFileOrNotify _
      (File:=File, SheetName:=Mobiles.ModelsTableName, ReadOnly:=ReadOnly)
  If Table Is Nothing Then
    Set tryBindModelsTable = Either.Create
    Exit Function
  End If
  
  With RecordListToTableBinder.Builder(Table)
    .WithKey "Size", 1
    If NameIsPrimaryKey Then .WithPrimaryKey "Name", 2 Else .WithKey "Name", 2
    .WithKey "Count", 3
    .WithMandatoryKey "File", 4
    .WithKey "SubCount1", Mobiles.SubColumn1
    .WithKey "SubCount2", Mobiles.SubColumn2
    .WithKey "SubCount3", Mobiles.SubColumn3
    .WithKey "SubCount4", Mobiles.SubColumn4
    .WithKey "ExcludeCount", 9
    .WithKey "ExcludeSubCount1", 10
    .WithKey "ExcludeSubCount2", 11
    .WithKey "ExcludeSubCount3", 12
    .WithKey "ExcludeSubCount4", 13
    Set tryBindModelsTable = Either.Create(.Build)
  End With
End Function

Public Function SecondaryModels _
                (ByVal SourceModels As IRecordList) As IRecordList
  If SourceModels.Count = 0 Then Exit Function
  Set SecondaryModels = SourceModels.GetCopy
  With SecondaryModels
    Dim i As Long
    For i = 1 To .Count
      .Record(i)("Count") = _
        CalcCount(.Record(i)("Count"), .Record(i)("ExcludeCount"))
      .Record(i)("SubCount1") = _
        CalcCount(.Record(i)("SubCount1"), .Record(i)("ExcludeSubCount1"))
      .Record(i)("SubCount2") = _
        CalcCount(.Record(i)("SubCount2"), .Record(i)("ExcludeSubCount2"))
      .Record(i)("SubCount3") = _
        CalcCount(.Record(i)("SubCount3"), .Record(i)("ExcludeSubCount3"))
      .Record(i)("SubCount4") = _
        CalcCount(.Record(i)("SubCount4"), .Record(i)("ExcludeSubCount4"))
    Next i
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
    .WithKey "Size", 2
    .WithKey "Line1", 3
    .WithKey "Line2", 4
    .WithKey "MaxBlocksInWidth", 5
    .WithKey "MaxPlacesInWidth", 6
    Set tryBindCategoriesTable = Either.Create(.Build)
  End With
End Function

Public Function PrimaryCategory _
                (ByVal SourceCategory As IRecord) As IRecord
  Set PrimaryCategory = SourceCategory.GetCopy
  PrimaryCategory!Line1 = _
    Mobiles.PrimaryCategoryCaptionPrefix & " — " & PrimaryCategory!Line1
End Function

Public Function SecondaryCategory _
                (ByVal SourceCategory As IRecord) As IRecord
  Set SecondaryCategory = SourceCategory.GetCopy
  SecondaryCategory!Name = _
    SecondaryCategory!Name & Mobiles.SecondaryCategoryNameSuffix
  SecondaryCategory!Line1 = _
    Mobiles.SecondaryCategoryCaptionPrefix & " — " & SecondaryCategory!Line1
End Function

Public Function tryBindAdditionalBlocksTable(ByVal File As IFileSpec) As IEither
  Dim Table As ITableFile
  Set Table = OpenTableFileOrNotify _
      (File:=File, SheetName:=Mobiles.AdditionalBlocksTableName, ReadOnly:=True)
  If Table Is Nothing Then
    Set tryBindAdditionalBlocksTable = Either.Create
    Exit Function
  End If
  With RecordListToTableBinder.Builder(Table)
    .WithPrimaryKey "Column", 1
    .WithKey "Name", 2
    Set tryBindAdditionalBlocksTable = Either.Create(.Build)
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

Public Sub ResetModelsCount(ByVal RecordList As IRecordList)
  Dim Record As IRecord
  For Each Record In RecordList.NewEnum
    Record.Field("Count") = 0
  Next Record
End Sub

Public Sub CountModelsInShapes(ByVal RecordList As IRecordList, _
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

Private Function CalcCount(ByVal BaseCount As String, _
                           ByVal MinusCount As String _
                           ) As String
  
  If IsFieldEmpty(MinusCount) Then Exit Function
  
  Dim MinusCountLng As Long
  MinusCountLng = ToLng(MinusCount)
  
  If MinusCountLng = 0 Then Exit Function
  
  Dim BaseCountLng As Long
  BaseCountLng = ToLng(BaseCount)
  
  Dim ResultLng As Long
  ResultLng = BaseCountLng - MinusCountLng
  If ResultLng < 0 Then ResultLng = 0
  CalcCount = VBA.CStr(ResultLng)

End Function

Public Sub CreateSheetOrNotify _
           (ByVal Category As IRecord, _
            ByVal Models As IRecordList, _
            ByVal AdditionalBlocks As IRecordList, _
            ByVal Sizes As IRecordList, _
            ByVal Log As Logger)
            
    With CategorySheet.CreateAndCompose( _
           Category:=Category, _
           Models:=Models, _
           AdditionalBlocks:=AdditionalBlocks, _
           Sizes:=Sizes _
         )
      If .IsError Then
        Log.Add "Мобайлов категории " & Category!Name & " не найдено"
      ElseIf .SuccessValue.FailedFiles.Count > 0 Then
        LogFailedFiles .SuccessValue.FailedFiles, Log
      End If
    End With

End Sub

Public Function CreateCaption(ByVal Caption As structCaption) As Shape
  With Caption
    Dim Text As String
    Text = .Line1
    If Not .Line2 = "" Then
      Text = Text & vbCrLf & _
             VBA.Replace(.Line2, Mobiles.CaptionsNewLineSymbol, vbCrLf)
    End If
    Set CreateCaption = ActiveLayer.CreateArtisticText( _
                          Left:=0, Bottom:=0, _
                          Text:=Text, _
                          Size:=.FontSize, _
                          Italic:=cdrTrue _
                        )
    If .Line1Bold Then
      CreateCaption.Text.Story.Lines(1).Style = cdrBoldItalicFontStyle
    End If
    CreateCaption.Fill.ApplyUniformFill CreateColor(.Color)
  End With
End Function

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

Public Sub LogFailedFiles(ByVal FailedFiles As IList, _
                          ByVal Log As Logger)
  Dim i As Long
  For i = 1 To FailedFiles.Count
    Log.Add "Не найден или не открывается файл " & FailedFiles(i)
  Next i
End Sub

Public Function IsFieldEmpty(ByVal Field As String) As Boolean
  IsFieldEmpty = True
  On Error GoTo ErrorHandler
  If Field = "" Then Exit Function
  If Field = "0" Then Exit Function
  If VBA.CStr(Field) = 0 Then Exit Function
ErrorHandler:
  On Error GoTo 0
  IsFieldEmpty = False
End Function

Private Function ToLng(ByVal Value As String) As Long
  If Value = "" Then Exit Function
  ToLng = VBA.CLng(Value)
End Function
