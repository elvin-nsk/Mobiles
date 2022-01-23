Attribute VB_Name = "Mobiles"
'===============================================================================
' ������           : Mobiles
' ������           : 2022.01.24
' ����             : https://github.com/elvin-nsk
' �����            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = True

'===============================================================================

Public Const ModelsTableName As String = "���-����"
Public Const CategoriesTableName As String = "���������"
Public Const AdditionalBlocksTableName As String = "����"
Public Const SizesTableName As String = "�������"

Public Const SizesDelimiterSymbol As String = ","
Public Const SizesMultiplierSymbol As String = "x"
Public Const CaptionsNewLineSymbol As String = ";"

Public Const SubColumn1 As Long = 5
Public Const SubColumn2 As Long = 6
Public Const SubColumn3 As Long = 7
Public Const SubColumn4 As Long = 8

Public Const CaptionsColor As String = "CMYK,USER,0,0,0,100"

Public Const DebugMobilesRootRepalceFrom As String = "C:\��\"
Public Const DebugMobilesRootRepalceTo As String = "e:\WORK\������� Corel\�� �����\������� �����\Mobiles\���������\2021-11-29\Data\"

'===============================================================================

Sub CountMobilesToTable()
  
  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then
    VBA.MsgBox "�������� �������"
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
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
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
  'Helpers.DebugPathsReplace Models
  
  Dim SecondaryModels As IRecordList
  Set SecondaryModels = Helpers.SecondaryModels(Models)
  
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
  
    Helpers.CreateSheetOrNotify _
              Category:=Helpers.PrimaryCategory(Category), _
              Models:=Models, _
              AdditionalBlocks:=AdditionalBlocks, _
              Sizes:=Sizes, _
              Log:=Log
              
    Helpers.CreateSheetOrNotify _
              Category:=Helpers.SecondaryCategory(Category), _
              Models:=SecondaryModels, _
              AdditionalBlocks:=AdditionalBlocks, _
              Sizes:=Sizes, _
              Log:=Log
    
    PBar.Update
    If PBar.Cancelled Then GoTo Finally
  
  Next Category
  
Finally:
  Optimization = False
  Application.Refresh
  Log.Check "����������"
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally

End Sub

'===============================================================================
' �����
'===============================================================================

Private Sub testWidth()
  ActiveDocument.Unit = cdrMillimeter
  ActivePage.SizeWidth = 45720 '������������ ������ ��������
  Debug.Print ConvertUnits(1800, cdrInch, ActiveDocument.Unit) '1800 ������
End Sub

Private Sub testRecordBuilder()
  Dim Rec As IRecord
  Dim RecFactory As IRecordFactory
  Set RecFactory = Record
  With RecFactory.Builder(StubKeys)
    .WithField "����1", 12
    .WithField "����2", "��������"
    Set Rec = .Build
  End With
  With Rec
    Debug.Assert .Field("����1") = 12
    Debug.Assert .Field("����2") = "��������"
    Debug.Assert .IsChanged = False
    .Field("����1") = 555
    .Field("����2") = "other"
    Debug.Assert .Field("����1") = 555
    Debug.Assert .Field("����2") = "other"
    Debug.Assert .IsChanged = True
    Dim RecCopy As IRecord
    Set RecCopy = .GetCopy
    Debug.Assert RecCopy.Field("����1") = 555
    Debug.Assert RecCopy.Field("����2") = "other"
    Debug.Assert RecCopy.IsChanged = False
  End With
End Sub

Private Sub testRecordList()
  With RecordList.Create(StubKeys)
    .BuildRecord.WithField("����1", 12).WithField("����2", "��������").Build
    .BuildRecord.WithField("����1", 55).WithField("����2", "Neo").Build
    Debug.Assert .Count = 2
    Debug.Assert .RecordExists(2) = True
    Debug.Assert .RecordExists(15) = False
    Debug.Assert .PrimaryKeySet = False
    Debug.Assert .Record(1).Field("����1") = 12
    Debug.Assert .Record(2)("����2") = "Neo"
    .Record(1).Field("����1") = 777
    Debug.Assert .Record(1).Field("����1") = 777
    Debug.Assert .Filter.Fields(777).Count = 1
    Debug.Assert .Filter.Fields("NoSuchValue").Count = 0
    Dim RecListCopy As IRecordList
    Set RecListCopy = .GetCopy
    Debug.Assert RecListCopy.Record(1).Field("����1") = 777
    Debug.Assert RecListCopy.Filter.Fields(777).Count = 1
    Debug.Assert RecListCopy.Filter.Fields("NoSuchValue").Count = 0
    RecListCopy.Record(2)("����2") = "NewForCopy"
    Debug.Assert .Record(2)("����2") = "Neo"
  End With
End Sub

Private Sub testRecordListWithPrimaryKey()
  With RecordList.Create(StubKeys, "����1")
    .BuildRecord.WithField("����1", "����").WithField("����2", "��������").Build
    .BuildRecord.WithField("����1", "����").WithField("����2", "Neo").Build
    .BuildRecord.WithField("����1", "����").WithField("����2", "Trinity").Build
    .BuildRecord.WithField("����1", "�������").WithField("����2", "Trinity").Build
    .BuildRecord.WithField("����1", "�����").WithField("����2", "").Build
    Debug.Assert .Count = 5
    Debug.Assert .RecordExists("����") = True
    Debug.Assert .RecordExists("��������") = False
    Debug.Assert .PrimaryKeySet = True
    Debug.Assert .PrimaryFieldExists("����") = True
    Debug.Assert .PrimaryFieldExists("����") = False
    Debug.Assert .Record("����")("����2") = "Neo"
    Debug.Assert .Record("����")("����2") = "Trinity"
    Debug.Assert .Record(1).ContainsLike("����*") = True
    Debug.Assert .Filter.FieldsLike("����*", "����1").Count = 2
    Debug.Assert .Filter.Fields(Array("��������", "Trinity"), "����2").Count = 3
    Debug.Assert .Filter.NotFields(Array("�������", "Trinity")).Count = 3
    Debug.Assert .Filter.NotFieldsEmpty("����2").Count = 4
    Dim RecListCopy As IRecordList
    Set RecListCopy = .GetCopy
    Debug.Assert RecListCopy.Filter.FieldsLike("����*", "����1").Count = 2
    Debug.Assert RecListCopy.Filter.Fields(Array("��������", "Trinity"), "����2").Count = 3
  End With
End Sub

Private Sub testBinder()
  Dim Table As ITableFile
  Dim File As IFileSpec
  Set File = FileSpec.Create("e:\WORK\������� Corel\�� �����\������� �����\Mobiles\���������\2021-11-16\test.xlsx")
  Set Table = ExcelConnection.Create(File, "���-����", 2)
  Dim Binder As IRecordListToTableBinder
  With RecordListToTableBinder.Builder(Table)
    .WithKey "Count", 3
    .WithPrimaryKey "Path", 4
    .WithUnboundKey "�������������� ����"
    Set Binder = .Build
  End With
  With Binder
    .RecordList(1)("�������������� ����") = "First"
    .RecordList(2)("�������������� ����") = 2#
    Debug.Print .RecordList(1)("�������������� ����")
    Debug.Print .RecordList(2)("�������������� ����")
    Debug.Print .RecordList.Count
    Debug.Print .RecordList(1)("Path")
    .RecordList(1)("Path") = "1111"
  End With
End Sub

Private Sub testExcelConnection()
  Dim Table As ITableFile
  Dim File As IFileSpec
  Set File = FileSpec.Create("e:\WORK\������� Corel\�� �����\������� �����\Mobiles\���������\2021-11-16\test.xlsx")
  With ExcelConnection.Create(File, "���-����")
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

  Const File = "e:\WORK\������� Corel\�� �����\������� �����\Mobiles\���������\2021-11-16\test.xlsx"
  
  Dim RecordSet As Object 'ADODB.RecordSet
  
  Dim Connection As Object 'ADODB.Connection
  Set Connection = VBA.CreateObject("ADODB.Connection")
  Dim SheetName As String
  With Connection
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Properties("Extended Properties").Value = "Excel 12.0;HDR=No"
    .Open File
    With .OpenSchema(adSchemaTables)
      SheetName = "���-����$" '.Fields("table_name").Value
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
          3, 0, 10, 5
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
    .Add "123" & vbCrLf & "����"
    .Check
  End With
End Sub

Private Sub testCreateCaption()
  With New structCaption
    .Line1 = "��������� �������"
    .Line2 = "������ �� ������� ������ + ������� ���������" & vbCrLf & _
             "+ �������� �� ��������� ������"
    .Color = CaptionsColor
    .FontSize = 50
    .Line1Bold = True
    Helpers.CreateCaption .Self
  End With
End Sub

Private Function StubKeys() As IList
  Set StubKeys = List.CreateFrom("����1", "����2", "3")
End Function
