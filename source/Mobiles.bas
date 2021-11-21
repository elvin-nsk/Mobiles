Attribute VB_Name = "Mobiles"
'===============================================================================
' Макрос           : Mobiles
' Версия           : 2021.11.21
' Сайт             : https://github.com/elvin-nsk
' Автор            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = True

'===============================================================================

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
      Set File = .Value
    End If
  End With
  
  Dim Table As ITableFile
  With Helpers.OpenTableFile(File)
    If .IsError Then
      VBA.MsgBox "Ошибка чтения файла", vbCritical
      Exit Sub
    Else
      Set Table = .Value
    End If
  End With
  
  Dim Mobiles As Dictionary
  Set Mobiles = Helpers.GetMobilesFromTable(Table)
  Helpers.CountMobilesInShapes Mobiles, ActiveSelectionRange
  Helpers.WriteMobileCountsToTable Mobiles, Table
  
  Set Table = Nothing
  
  lib_elvin.BoostStart , RELEASE
  
  With CompositeSheet.Create(Mobiles)
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

Private Sub testExcelNew()
  Dim WB As Excel.Workbook
  Set WB = Excel.Application.Workbooks.Add
  WB.ActiveSheet.Cells(1, 2) = "123"
  WB.SaveAs "e:\temp\123.xlsx"
  WB.Close
End Sub

Private Sub testExcelEdit()
  Dim WB As Excel.Workbook
  Set WB = Excel.Application.Workbooks.Open("e:\temp\123.xlsx")
  WB.ActiveSheet.Cells(1, 1) = "Первая ячейка"
  WB.Save
  WB.Close
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

Private Sub testSaveAs()
  Dim FileName As Variant
  FileName = Excel.Application.GetSaveAsFilename("e:\temp\new file.xlsx", "Excel Files (*.xlsx), *.xlsx")
  Debug.Print FileName
End Sub

Private Sub testLoad()
  Dim FileName As Variant
  FileName = Excel.Application.GetOpenFileName("Excel Files (*.xlsx), *.xlsx")
  Debug.Print FileName
End Sub

Private Sub testEither()
  Dim Result As IEither
  Set Result = Either.Create()
  If Result.IsError Then
    Debug.Print "Error"
  Else
    Debug.Print "Success"
  End If
  Debug.Print VBA.CLng(True)
End Sub
