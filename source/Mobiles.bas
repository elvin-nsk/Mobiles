Attribute VB_Name = "Mobiles"
'===============================================================================
' ������           : Mobiles
' ������           : 2022.01.05
' ����             : https://github.com/elvin-nsk
' �����            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = False

'===============================================================================

Public Const MainTableName As String = "���-����"
Public Const CategoriesTableName As String = "���������"
Public Const SubTableName As String = "����"
Public Const SizesTableName As String = "�������"

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
  With Helpers.tryBindMainTable _
               (File:=File, NameIsPrimaryKey:=True, ReadOnly:=False)
    If .IsError Then Exit Sub
    Set Binder = .SuccessValue
  End With
   
  Helpers.ResetMobilesCount Binder.RecordList
  Helpers.CountMobilesInShapes Binder.RecordList, ActiveSelectionRange
  
Finally:
  Set Binder = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally

End Sub

Sub CreateSheetsFromTable()

  If RELEASE Then On Error GoTo Catch
  
  Dim File As IFileSpec
  With Helpers.tryGetExcelFile
    If .IsError Then Exit Sub
    Set File = .SuccessValue
  End With
  
  Dim Binder As IRecordListToTableBinder
  With Helpers.tryBindMainTable _
               (File:=File, NameIsPrimaryKey:=False, ReadOnly:=True)
    If .IsError Then Exit Sub
    Set Binder = .SuccessValue
  End With
  #If Not RELEASE Then
    Helpers.DebugPathsReplace Binder.RecordList
  #End If
  
  'Helpers.ValidateMainTable Binder.RecordList
  
  lib_elvin.BoostStart , RELEASE
  
  'With CompositeSheet.Create(Binder.RecordList, RELEASE)
  '  If .FailedFiles.Count > 0 Then
  '    Helpers.Report .FailedFiles
  '  End If
  'End With
    
  Debug.Print Binder.RecordList.Count
  Debug.Assert Binder.RecordList(5)("File") <> ""
  
Finally:
  Set Binder = Nothing
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally

End Sub

'===============================================================================
' �����
'===============================================================================

Private Sub testExcelEditLateBinding()
  Dim App As Object
  Set App = VBA.CreateObject("Excel.Application")
  Dim WB As Object
  Set WB = App.Workbooks.Open("e:\temp\123.xlsx")
  WB.ActiveSheet.Cells(1, 1) = "Late"
  WB.Save
  WB.Close
End Sub

Private Sub testWidth()
  ActiveDocument.Unit = cdrMillimeter
  ActivePage.SizeWidth = 30000
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
  End With
End Sub

Private Sub testRecordListWithPrimaryKey()
  With RecordList.Create(StubKeys, "����1")
    .BuildRecord.WithField("����1", "����").WithField("����2", "��������").Build
    .BuildRecord.WithField("����1", "����").WithField("����2", "Neo").Build
    .BuildRecord.WithField("����1", "����").WithField("����2", "Trinity").Build
    .BuildRecord.WithField("����1", "�������").WithField("����2", "Trinity").Build
    Debug.Assert .Count = 4
    Debug.Assert .RecordExists("����") = True
    Debug.Assert .RecordExists("��������") = False
    Debug.Assert .PrimaryKeySet = True
    Debug.Assert .PrimaryFieldExists("����") = True
    Debug.Assert .PrimaryFieldExists("����") = False
    Debug.Assert .Record("����")("����2") = "Neo"
    Debug.Assert .Record("����")("����2") = "Trinity"
    Debug.Assert .Record(1).ContainsLike("����*") = True
    Debug.Assert .Filter.FieldsLike("����*", "����1").Count = 2
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

Sub testADODB()

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

Private Function StubKeys() As Collection
  Set StubKeys = New Collection
  StubKeys.Add "����1"
  StubKeys.Add "����2"
  StubKeys.Add "3"
End Function
