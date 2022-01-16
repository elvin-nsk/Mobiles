Attribute VB_Name = "Mobiles"
'===============================================================================
' Макрос           : Mobiles
' Версия           : 2022.01.17
' Сайт             : https://github.com/elvin-nsk
' Автор            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = True

'===============================================================================

Public Const ModelsTableName As String = "Чек-лист"
Public Const CategoriesTableName As String = "Категории"
Public Const AdditionalBlocksTableName As String = "Виды"
Public Const SizesTableName As String = "Размеры"

Public Const SizesDelimiterSymbol As String = ","
Public Const SizesMultiplierSymbol As String = "x"
Public Const CaptionsNewLineSymbol As String = ";"

Public Const SubColumn1 As Long = 5
Public Const SubColumn2 As Long = 6
Public Const SubColumn3 As Long = 7
Public Const SubColumn4 As Long = 8

Public Const CaptionsColor As String = "CMYK,USER,0,0,0,100"

Public Const DebugMobilesRootRepalceFrom As String = "C:\МК\"
Public Const DebugMobilesRootRepalceTo As String = "e:\WORK\макросы Corel\на заказ\Дмитрий Шмыга\Mobiles\материалы\2021-11-29\Data\"

'===============================================================================

Sub CountMobilesToTable()
  
  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then
    VBA.MsgBox "Выберите мобайлы"
    Exit Sub
  End If
  
  Dim File As IFileSpec
  With Helpers.tryGetExcelFile
    If .IsError Then Exit Sub
    Set File = .SuccessValue
  End With
  
  Dim Binder As IRecordListToTableBinder
  With Helpers.tryBindModelsTable _
               (File:=File, NameIsPrimaryKey:=True, ReadOnly:=False)
    If .IsError Then Exit Sub
    Set Binder = .SuccessValue
  End With
   
  Helpers.ResetModelsCount Binder.RecordList
  Helpers.CountModelsInShapes Binder.RecordList, ActiveSelectionRange
  
Finally:
  Set Binder = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

Sub CreateSheetsFromTable()

  If RELEASE Then On Error GoTo Catch
  
  Dim Log As New Logger
  
  Dim File As IFileSpec
  With Helpers.tryGetExcelFile
    If .IsError Then Exit Sub
    Set File = .SuccessValue
  End With
  
  Dim Models As IRecordList
  With Helpers.tryBindModelsTable _
               (File:=File, NameIsPrimaryKey:=False, ReadOnly:=True)
    If .IsError Then Exit Sub
    Set Models = .SuccessValue.RecordList
  End With
  '#If Not RELEASE Then
    Helpers.DebugPathsReplace Models
  '#End If
  
  Dim Categories As IRecordList
  With Helpers.tryBindCategoriesTable(File)
    If .IsError Then Exit Sub
    Set Categories = .SuccessValue.RecordList
  End With
  
  Dim AdditionalBlocks As IRecordList
  With Helpers.tryBindAdditionalBlocksTable(File)
    If .IsError Then Exit Sub
    Set AdditionalBlocks = .SuccessValue.RecordList
  End With
  
  Dim Sizes As IRecordList
  With Helpers.tryBindSizesTable(File)
    If .IsError Then Exit Sub
    Set Sizes = .SuccessValue.RecordList
  End With
  
  Optimization = RELEASE
  
  Dim PBar As IProgressBar
  Set PBar = ProgressBar.CreateNumeric(Categories.Count)
  PBar.Cancellable = True
  
  Dim Category As IRecord
  For Each Category In Categories.NewEnum
  
    With CategorySheet.CreateAndCompose( _
           Category:=Category, _
           Models:=Models, _
           AdditionalBlocks:=AdditionalBlocks, _
           Sizes:=Sizes _
         )
      If .IsError Then
        Log.Add "Мобайлов категории " & Category!Name & " не найдено"
        Exit Sub
      End If
      With .SuccessValue
        If .FailedFiles.Count > 0 Then
          Helpers.LogFailedFiles .FailedFiles, Log
        End If
      End With
    End With
    PBar.Update
    If PBar.Cancelled Then GoTo Finally
  
  Next Category
  
Finally:
  Optimization = False
  Application.Refresh
  Log.Check "Информация"
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

'===============================================================================
' тесты
'===============================================================================

Private Sub testWidth()
  ActiveDocument.Unit = cdrMillimeter
  ActivePage.SizeWidth = 45720 'максимальный размер страницы
  Debug.Print ConvertUnits(1800, cdrInch, ActiveDocument.Unit) '1800 дюймов
End Sub

Private Sub testRecordBuilder()
  Dim Rec As IRecord
  Dim RecFactory As IRecordFactory
  Set RecFactory = Record
  With RecFactory.Builder(StubKeys)
    .WithField "Поле1", 12
    .WithField "Поле2", "Значение"
    Set Rec = .Build
  End With
  With Rec
    Debug.Assert .Field("Поле1") = 12
    Debug.Assert .Field("Поле2") = "Значение"
    Debug.Assert .IsChanged = False
    .Field("Поле1") = 555
    .Field("Поле2") = "other"
    Debug.Assert .Field("Поле1") = 555
    Debug.Assert .Field("Поле2") = "other"
    Debug.Assert .IsChanged = True
  End With
End Sub

Private Sub testRecordList()
  With RecordList.Create(StubKeys)
    .BuildRecord.WithField("Поле1", 12).WithField("Поле2", "Значение").Build
    .BuildRecord.WithField("Поле1", 55).WithField("Поле2", "Neo").Build
    Debug.Assert .Count = 2
    Debug.Assert .RecordExists(2) = True
    Debug.Assert .RecordExists(15) = False
    Debug.Assert .PrimaryKeySet = False
    Debug.Assert .Record(1).Field("Поле1") = 12
    Debug.Assert .Record(2)("Поле2") = "Neo"
    .Record(1).Field("Поле1") = 777
    Debug.Assert .Record(1).Field("Поле1") = 777
    Debug.Assert .Filter.Fields(777).Count = 1
    Debug.Assert .Filter.Fields("NoSuchValue").Count = 0
  End With
End Sub

Private Sub testRecordListWithPrimaryKey()
  With RecordList.Create(StubKeys, "Поле1")
    .BuildRecord.WithField("Поле1", "Вася").WithField("Поле2", "Значение").Build
    .BuildRecord.WithField("Поле1", "Петя").WithField("Поле2", "Neo").Build
    .BuildRecord.WithField("Поле1", "Джон").WithField("Поле2", "Trinity").Build
    .BuildRecord.WithField("Поле1", "Джонсон").WithField("Поле2", "Trinity").Build
    .BuildRecord.WithField("Поле1", "Серёжа").WithField("Поле2", "").Build
    Debug.Assert .Count = 5
    Debug.Assert .RecordExists("Джон") = True
    Debug.Assert .RecordExists("Хамелеон") = False
    Debug.Assert .PrimaryKeySet = True
    Debug.Assert .PrimaryFieldExists("Джон") = True
    Debug.Assert .PrimaryFieldExists("Зязя") = False
    Debug.Assert .Record("Петя")("Поле2") = "Neo"
    Debug.Assert .Record("Джон")("Поле2") = "Trinity"
    Debug.Assert .Record(1).ContainsLike("Знач*") = True
    Debug.Assert .Filter.FieldsLike("Джон*", "Поле1").Count = 2
    Debug.Assert .Filter.Fields(Array("Значение", "Trinity"), "Поле2").Count = 3
    Debug.Assert .Filter.NotFields(Array("Джонсон", "Trinity")).Count = 3
    Debug.Assert .Filter.NotFieldsEmpty("Поле2").Count = 4
  End With
End Sub

Private Sub testBinder()
  Dim Table As ITableFile
  Dim File As IFileSpec
  Set File = FileSpec.Create("e:\WORK\макросы Corel\на заказ\Дмитрий Шмыга\Mobiles\материалы\2021-11-16\test.xlsx")
  Set Table = ExcelConnection.Create(File, "Чек-лист", 2)
  Dim Binder As IRecordListToTableBinder
  With RecordListToTableBinder.Builder(Table)
    .WithKey "Count", 3
    .WithPrimaryKey "Path", 4
    .WithUnboundKey "Дополнительное поле"
    Set Binder = .Build
  End With
  With Binder
    .RecordList(1)("Дополнительное поле") = "First"
    .RecordList(2)("Дополнительное поле") = 2#
    Debug.Print .RecordList(1)("Дополнительное поле")
    Debug.Print .RecordList(2)("Дополнительное поле")
    Debug.Print .RecordList.Count
    Debug.Print .RecordList(1)("Path")
    .RecordList(1)("Path") = "1111"
  End With
End Sub

Private Sub testExcelConnection()
  Dim Table As ITableFile
  Dim File As IFileSpec
  Set File = FileSpec.Create("e:\WORK\макросы Corel\на заказ\Дмитрий Шмыга\Mobiles\материалы\2021-11-16\test.xlsx")
  With ExcelConnection.Create(File, "Чек-лист")
    Debug.Print .Cell(1, 1)
    Debug.Print .Cell(19, 3)
    Debug.Print .Cell(2, 1)
    .Cell(2, 1) = lib_elvin.RndInt(1, 100)
    Debug.Print .Cell(2, 1)
    '.ForceSave
    '.ForceClose
  End With
  
End Sub

Private Sub testADODB()

  Const adLockOptimistic = 3
  Const adLockReadOnly = 1
  Const adUseServer = 2
  Const adUseClient = 3
  Const adSchemaTables = 20

  Const File = "e:\WORK\макросы Corel\на заказ\Дмитрий Шмыга\Mobiles\материалы\2021-11-16\test.xlsx"
  
  Dim RecordSet As Object 'ADODB.RecordSet
  
  Dim Connection As Object 'ADODB.Connection
  Set Connection = VBA.CreateObject("ADODB.Connection")
  Dim SheetName As String
  With Connection
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Properties("Extended Properties").Value = "Excel 12.0;HDR=No"
    .Open File
    With .OpenSchema(adSchemaTables)
      SheetName = "Чек-лист$" '.Fields("table_name").Value
      .Close
    End With
    'Debug.Print SheetName
    Set RecordSet = VBA.CreateObject("ADODB.RecordSet")
    With RecordSet
      .ActiveConnection = Connection
      .LockType = adLockOptimistic
      .CursorLocation = adUseServer
      .Source = "Select * from [" & SheetName & "]"
      .Open
      Debug.Print .RecordCount
      .MoveFirst
      .Move 1
      .Fields(0) = "123"
      .Update 'save
      '.CancelUpdate
      .Close
    End With
  End With
  
  Connection.Close

End Sub

Private Sub testBlock()
  
  Dim Models As IList
  With List.Create
    Set Models = .Self
  End With
  
  Dim Caption As New structCaption
  With Caption
  End With
  
  With CategorySheetBlock.Create(Models, Caption, _
                                 10, 0, _
                                 50, 50, _
                                 FreePoint.Create(0, 0))
  End With
End Sub

Private Sub testStackable()
  With Stackable.Create(ActiveLayer.Shapes.First)
    .PivotX = 10
    .PivotY = 5
    Debug.Assert .PivotX = 10
    Debug.Assert .PivotY = 5
    Debug.Print .Width
    Debug.Print .Height
  End With
End Sub

Private Sub testStacker()
  ActiveDocument.Unit = cdrMillimeter
  Dim Stackables As New Collection
  Dim Shape As Shape
  For Each Shape In ActiveLayer.Shapes
    Stackables.Add Stackable.Create(Shape)
  Next Shape
  Stacker.CreateAndStack Stackables, FreePoint.Create(0, ActivePage.TopY), _
          , 0, 10, 5
End Sub

Private Sub testList()
  With List.Create
    .Add "123"
    Debug.Assert .Item(1) = "123"
  End With
  With List.CreateFrom(1, 2, 3)
    Debug.Assert .Item(2) = 2
  End With
End Sub

Private Sub testLogger()
  With New Logger
    .Add "123" & vbCrLf & "Вася"
    .Check
  End With
End Sub

Private Sub testCreateCaption()
  With New structCaption
    .Line1 = "Магнитные мобайлы"
    .Line2 = "Печать на матовой пленке + матовая ламинация" & vbCrLf & _
             "+ накатать на магнитную основу"
    .Color = CaptionsColor
    .FontSize = 50
    .Line1Bold = True
    Helpers.CreateCaption .Self
  End With
End Sub

Private Function StubKeys() As IList
  Set StubKeys = List.CreateFrom("Поле1", "Поле2", "3")
End Function
