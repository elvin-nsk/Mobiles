Attribute VB_Name = "Mobiles"
'===============================================================================
' Макрос           : Mobiles
' Версия           : 2022.01.02
' Сайт             : https://github.com/elvin-nsk
' Автор            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = False

'===============================================================================

'Public Const DebugMobilesRootRepalceFrom As String = "C:\МК\"
'Public Const DebugMobilesRootRepalceTo As String = "e:\WORK\макросы Corel\на заказ\Дмитрий Шмыга\Mobiles\материалы\2021-11-29\Data\"
Public Const DebugMobilesRootRepalceFrom As String = "C:\МобайлыМакеты\"
Public Const DebugMobilesRootRepalceTo As String = "e:\WORK\макросы Corel\на заказ\Дмитрий Шмыга\Mobiles\материалы\2021-11-16\Data\МобайлыМакеты\"


'===============================================================================

Sub CountMobilesToTable()

  Dim Table As ITableFile
  
  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then
    VBA.MsgBox "Выберите мобайлы"
    Exit Sub
  End If
  
  Dim File As IFileSpec
  With Helpers.GetExcelFile
    If .IsError Then Exit Sub
    Set File = .SuccessValue
  End With
  
  With Helpers.OpenTableFile(File:=File, ReadOnly:=False)
    If .IsError Then
      VBA.MsgBox "Ошибка чтения файла", vbCritical
      Exit Sub
    End If
    Set Table = .SuccessValue
  End With
  
  Dim Binder As IRecordListToTableBinder
  Set Binder = Helpers.BindMainTable(Table, "Name")
  
  Helpers.ResetMobilesCount Binder.RecordList
  Helpers.CountMobilesInShapes Binder.RecordList, ActiveSelectionRange
  
Finally:
  Set Table = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

Sub CreateSheetsFromTable()

  Dim Table As ITableFile

  If RELEASE Then On Error GoTo Catch
  
  Dim File As IFileSpec
  With Helpers.GetExcelFile
    If .IsError Then Exit Sub
    Set File = .SuccessValue
  End With
  
  With Helpers.OpenTableFile(File:=File, ReadOnly:=True)
    If .IsError Then
      VBA.MsgBox "Ошибка чтения файла", vbCritical
      Exit Sub
    End If
    Set Table = .SuccessValue
  End With
  
  Dim Binder As IRecordListToTableBinder
  Set Binder = Helpers.BindMainTable(Table, "File")
  
  lib_elvin.BoostStart , RELEASE
  
  With CompositeSheet.Create(Binder.RecordList, RELEASE)
    If .FailedFiles.Count > 0 Then
      Helpers.Report .FailedFiles
    End If
  End With
  
Finally:
  Set Table = Nothing
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

Sub Start()

  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then
    VBA.MsgBox "Выберите мобайлы"
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
      VBA.MsgBox "Ошибка чтения файла", vbCritical
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
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  If Not Table Is Nothing Then Table.ForceClose
  Resume Finally

End Sub

'===============================================================================
' тесты
'===============================================================================

Private Sub testExcelConnection()
  Dim Table As ITableFile
  Dim File As IFileSpec
  Set File = FileSpec.Create("e:\WORK\макросы Corel\на заказ\Дмитрий Шмыга\Mobiles\материалы\2021-11-16\test.xlsx")
  With ExcelConnection.CreateReadOnly(File, "Лист1")
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
  With RecordList.Create(MockFieldNames)
    .BuildRecord.WithField("Поле1", 12).WithField("Поле2", "Значение").Build
    .BuildRecord.WithField("Поле1", 55).WithField("Поле2", "Neo").Build
    Debug.Assert .Count = 2
    Debug.Assert .RecordExists(2) = True
    Debug.Assert .RecordExists(15) = False
    Debug.Assert .KeyFieldSet = False
    Debug.Assert .Record(1).Field("Поле1") = 12
    Debug.Assert .Record(2)("Поле2") = "Neo"
    .Record(1).Field("Поле1") = 777
    Debug.Assert .Record(1).Field("Поле1") = 777
  End With
End Sub

Private Sub testRecordListWithKeyField()
  With RecordList.Create(MockFieldNames, "Поле1")
    .BuildRecord.WithField("Поле1", "Вася").WithField("Поле2", "Значение").Build
    .BuildRecord.WithField("Поле1", "Петя").WithField("Поле2", "Neo").Build
    .BuildRecord.WithField("Поле1", "Джон").WithField("Поле2", "Trinity").Build
    .BuildRecord.WithField("Поле1", "Джонсон").WithField("Поле2", "Trinity").Build
    Debug.Assert .Count = 4
    Debug.Assert .RecordExists("Джон") = True
    Debug.Assert .RecordExists("Хамелеон") = False
    Debug.Assert .KeyFieldSet = True
    Debug.Assert .KeyFieldExists("Джон") = True
    Debug.Assert .KeyFieldExists("Зязя") = False
    Debug.Assert .Record("Петя")("Поле2") = "Neo"
    Debug.Assert .Record("Джон")("Поле2") = "Trinity"
    Debug.Assert .Record(1).ContainsLike("Знач*") = True
    Debug.Assert .FilterLike("Джон*").Count = 2
  End With
End Sub

Private Sub testBinder()
  Dim Table As ITableFile
  Dim File As IFileSpec
  Set File = FileSpec.Create("e:\WORK\макросы Corel\на заказ\Дмитрий Шмыга\Mobiles\материалы\2021-11-16\test.xlsx")
  Set Table = ExcelConnection.CreateReadOnly(File, "Лист1", 2)
  Dim Binder As IRecordListToTableBinder
  With RecordListToTableBinder.Builder(Table)
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
  MockFieldNames.Add "Поле1"
  MockFieldNames.Add "Поле2"
  MockFieldNames.Add "3"
End Function
