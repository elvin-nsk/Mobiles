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

#If VBA7 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal win As Long) As Long
#Else
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal win As Long) As Long
#End If

Private Sub Class_Initialize()
  If Not Me Is Helpers Then Err.Raise 425 'Invalid object use
End Sub

'===============================================================================

Public Function GetExcelFile() As IEither
  Const AppName As String = "Mobiles"
  Const Section As String = "Settings"
  Const Key As String = "TableLastPath"
  Dim Path As String
  Path = GetSetting _
    (AppName, Section, Key, "%USERPROFILE%\Desktop")
  Dim Result As Variant
  With Excel.Application.FileDialog(1)
    .ButtonName = "&Открыть"
    .InitialFilename = Path
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx", 1
    .Title = "Выберите таблицу категорий"
    .AllowMultiSelect = False
    SetForegroundWindow Excel.Application.Hwnd
    If .Show = -1 Then
      Set GetExcelFile = Either.Create(FileSpec.Create(.SelectedItems(1)))
      Path = GetExcelFile.Value.Path
      SaveSetting AppName, Section, Key, Path
    Else
      Set GetExcelFile = Either.Create()
    End If
  End With
End Function

Public Function OpenTableFile(ByVal File As IFileSpec) As IEither
  Dim Result As ITableFile
  On Error GoTo ErrorHandler
  Set Result = ExcelFile.Create(FileToBind:=File, StartingRow:=2)
  Set OpenTableFile = Either.Create(Result)
  Exit Function
ErrorHandler:
  Set OpenTableFile = Either.Create()
End Function

Public Function GetMobilesFromTable(ByVal Table As ITableFile) As Dictionary
  Dim Mobiles As Dictionary
  Set Mobiles = New Dictionary
  Dim Mobile As structMobile
  Dim Row As Long
  Dim Name As String
  Do
    Row = Row + 1
    Name = Table.Cell(Row, 1)
    tryAddMobile Table, Row, Mobiles
    If Row = 1000 Then Exit Do
  Loop Until Name = ""
  Set GetMobilesFromTable = Mobiles
End Function

Private Sub tryAddMobile(ByVal Table As ITableFile, _
                         ByVal Row As Long, _
                         ByVal Mobiles As Dictionary)
  
  Dim Name As String
  Name = Table.Cell(Row, 1)
  If Name = "" Then Exit Sub
  
  Dim Spec As String
  Spec = Table.Cell(Row, 3)
  If Spec = "" Then Exit Sub
  
  Dim Mobile As structMobile
  Set Mobile = New structMobile
  Set Mobile.File = FileSpec.Create(Spec)
  Mobiles.Add Name, Mobile

End Sub

Public Sub CountMobilesInShapes(ByVal Mobiles As Dictionary, _
                                ByVal Shapes As ShapeRange)
  Dim Name As String
  Dim Shape As Shape
  For Each Shape In Shapes
    Name = Shape.Name
    If Not Name = "" Then
      If Mobiles.Exists(Name) Then
        Mobiles(Name).Count = Mobiles(Name).Count + 1
      End If
    End If
  Next Shape
End Sub

Public Sub WriteMobileCountsToTable(ByVal Mobiles As Dictionary, _
                                    ByVal Table As ITableFile)
  Dim Row As Long
  Dim Name As String
  Do
    Row = Row + 1
    Name = Table.Cell(Row, 1)
    If Mobiles.Exists(Name) Then
      Table.Cell(Row, 2) = VBA.CStr(Mobiles.Item(Name).Count)
    End If
    If Row = 1000 Then Exit Do
  Loop Until Name = ""
End Sub

Public Sub Report(ByVal FailedFiles As Collection)
  Dim Text As String
  Text = "Следующие файлы не найдены:" & vbCrLf
  Dim i As Long
  For i = 1 To FailedFiles.Count
    Text = Text & FailedFiles(i) & vbCrLf
  Next i
  VBA.MsgBox Text, vbInformation, "Ошибка импорта"
End Sub
