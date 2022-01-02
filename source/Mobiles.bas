Attribute VB_Name = "Mobiles"
'===============================================================================
' ������           : MobilesDic
' ������           : 2021.11.23
' ����             : https://github.com/elvin-nsk
' �����            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = False

'===============================================================================

'Public Const DebugMobilesRootRepalceFrom As String = "C:\��\"
'Public Const DebugMobilesRootRepalceTo As String = "e:\WORK\������� Corel\�� �����\������� �����\Mobiles\���������\2021-11-29\Data\"
Public Const DebugMobilesRootRepalceFrom As String = "C:\�������������\"
Public Const DebugMobilesRootRepalceTo As String = "e:\WORK\������� Corel\�� �����\������� �����\Mobiles\���������\2021-11-16\Data\�������������\"


'===============================================================================

Sub CountMobilesToTable()

  Dim Table As ITableFile
  
  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then
    VBA.MsgBox "�������� �������"
    Exit Sub
  End If
  
  Dim File As IFileSpec
  With Helpers.GetExcelFile
    If .IsError Then
      Exit Sub
    Else
      Set File = .SuccessValue
    End If
  End With
  
  With Helpers.OpenTableFile(File)
    If .IsError Then
      VBA.MsgBox "������ ������ �����", vbCritical
      Exit Sub
    Else
      Set Table = .SuccessValue
    End If
  End With
  
  Dim Binder As IRecordSetToTableBinder
  Set Binder = Helpers.BindMainTable(Table, "Name")
  
  Helpers.ResetMobilesCount Binder.RecordSet
  Helpers.CountMobilesInShapes Binder.RecordSet, ActiveSelectionRange
  
Finally:
  Set Table = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally

End Sub

Sub CreateSheetsFromTable()

  Dim Table As ITableFile

  If RELEASE Then On Error GoTo Catch
  
  Dim File As IFileSpec
  With Helpers.GetExcelFile
    If .IsError Then
      Exit Sub
    Else
      Set File = .SuccessValue
    End If
  End With
  
  With Helpers.OpenTableFile(File:=File, ReadOnly:=True)
    If .IsError Then
      VBA.MsgBox "������ ������ �����", vbCritical
      Exit Sub
    Else
      Set Table = .SuccessValue
    End If
  End With
  
  Dim Binder As IRecordSetToTableBinder
  Set Binder = Helpers.BindMainTable(Table, "File")
  
  lib_elvin.BoostStart , RELEASE
  
  With CompositeSheet.Create(Binder.RecordSet, RELEASE)
    If .FailedFiles.Count > 0 Then
      Helpers.Report .FailedFiles
    End If
  End With
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  If Not Table Is Nothing Then Table.ForceClose
  Resume Finally

End Sub

Sub Start()

  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then
    VBA.MsgBox "�������� �������"
    Exit Sub
  End If
  
  Dim File As IFileSpec
  With Helpers.GetExcelFile
    If .IsError Then
      Exit Sub
    Else
      Set File = .SuccessValue
    End If
  End With
  
  Dim Table As ITableFile
  With Helpers.OpenTableFile(File)
    If .IsError Then
      VBA.MsgBox "������ ������ �����", vbCritical
      Exit Sub
    Else
      Set Table = .SuccessValue
    End If
  End With
  
  Dim MobilesDic As Dictionary
  Set MobilesDic = Helpers.GetMobilesFromTable(Table)
  Helpers.CountMobilesInShapes MobilesDic, ActiveSelectionRange
  Helpers.WriteMobileCountsToTable MobilesDic, Table
  
  Set Table = Nothing
  
  lib_elvin.BoostStart , RELEASE
  
  With CompositeSheet.Create(MobilesDic, RELEASE)
    If .FailedFiles.Count > 0 Then
      Helpers.Report .FailedFiles
    End If
  End With
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  If Not Table Is Nothing Then Table.ForceClose
  Resume Finally

End Sub

'===============================================================================
' �����
'===============================================================================

Private Sub testExcelConnection()
  Dim Table As ITableFile
  Dim File As IFileSpec
  Set File = FileSpec.Create("e:\WORK\������� Corel\�� �����\������� �����\Mobiles\���������\2021-11-16\test.xlsx")
  With ExcelConnection.CreateReadOnly(File, "����1")
    Debug.Print .Cell(1, 1)
    Debug.Print .Cell(19, 3)
    Debug.Print .Cell(1, 2)
    '.Cell(2, 3) = 123
  End With
  
End Sub

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
  With RecFactory.Builder(MockFieldNames)
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

Private Sub testRecordSet()
  Dim RecSet As IRecordSet
  With RecordSet.Create(MockFieldNames)
    .BuildRecord.WithField("����1", 12).WithField("����2", "��������").Build
    .BuildRecord.WithField("����1", 55).WithField("����2", "Neo").Build
    Debug.Assert .Count = 2
    Debug.Assert .RecordExists(2) = True
    Debug.Assert .RecordExists(15) = False
    Debug.Assert .KeyFieldSet = False
    Debug.Assert .Record(1).Field("����1") = 12
    Debug.Assert .Record(2)("����2") = "Neo"
    .Record(1).Field("����1") = 777
    Debug.Assert .Record(1).Field("����1") = 777
  End With
End Sub

Private Sub testRecordSetWithKeyField()
  Dim RecSet As IRecordSet
  With RecordSet.Create(MockFieldNames, "����1")
    .BuildRecord.WithField("����1", "����").WithField("����2", "��������").Build
    .BuildRecord.WithField("����1", "����").WithField("����2", "Neo").Build
    .BuildRecord.WithField("����1", "����").WithField("����2", "Trinity").Build
    .BuildRecord.WithField("����1", "�������").WithField("����2", "Trinity").Build
    Debug.Assert .Count = 4
    Debug.Assert .RecordExists("����") = True
    Debug.Assert .RecordExists("��������") = False
    Debug.Assert .KeyFieldSet = True
    Debug.Assert .KeyFieldExists("����") = True
    Debug.Assert .KeyFieldExists("����") = False
    Debug.Assert .Record("����")("����2") = "Neo"
    Debug.Assert .Record("����")("����2") = "Trinity"
    Debug.Assert .Record(1).ContainsLike("����*") = True
    Debug.Assert .FilterLike("����*").Count = 2
  End With
End Sub

Private Sub testBinder()
  Dim Table As ITableFile
  Dim File As IFileSpec
  Set File = FileSpec.Create("e:\WORK\������� Corel\�� �����\������� �����\Mobiles\���������\2021-11-16\test.xlsx")
  Set Table = ExcelConnection.CreateReadOnly(File, "����1", 2)
  Dim Binder As IRecordSetToTableBinder
  With RecordSetToTableBinder.Builder(Table)
    .WithField "Count", 2
    .WithField "Path", 3
    .WithKeyField "Path"
    Set Binder = .Build
  End With
  With Binder
    Debug.Print .RecordSet.Count
    Debug.Print .RecordSet(1)("Path")
    '.RecordSet(1)("Path") = "1111"
  End With
End Sub

Private Function MockFieldNames() As Collection
  Set MockFieldNames = New Collection
  MockFieldNames.Add "����1"
  MockFieldNames.Add "����2"
  MockFieldNames.Add "3"
End Function
