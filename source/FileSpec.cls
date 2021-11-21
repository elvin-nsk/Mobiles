VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IFileSpec

'===============================================================================

Private Type typeThis
  Path As String
  NameWithoutExt As String
  Ext As String
End Type
Private This As typeThis

'===============================================================================

Public Function Create(Optional ByVal File As String) As IFileSpec
  With New FileSpec
    Set Create = .Init(File)
  End With
End Function

Friend Function Init(ByVal File As String) As IFileSpec
  With This
    .Path = GetFilePath(File)
    .NameWithoutExt = GetFileNameWithoutExt(GetFileName(File))
    .Ext = GetFileExt(GetFileName(File))
  End With
  Set Init = Me
End Function

'===============================================================================

Private Property Get IFileSpec_Ext() As String
  IFileSpec_Ext = This.Ext
End Property
Private Property Let IFileSpec_Ext(ByVal Value As String)
  This.Ext = Value
End Property

Private Function IFileSpec_ToString() As String
  IFileSpec_ToString = This.Path & This.NameWithoutExt & AddDotIfNotEmpty(This.Ext)
End Function

Private Property Get IFileSpec_Name() As String
  IFileSpec_Name = This.NameWithoutExt & AddDotIfNotEmpty(This.Ext)
End Property
Private Property Let IFileSpec_Name(ByVal Value As String)
  With This
    .NameWithoutExt = GetFileNameWithoutExt(Value)
    .Ext = GetFileExt(Value)
  End With
End Property

Private Property Get IFileSpec_NameWithoutExt() As String
  IFileSpec_NameWithoutExt = This.NameWithoutExt
End Property
Private Property Let IFileSpec_NameWithoutExt(ByVal Value As String)
  This.NameWithoutExt = Value
End Property

Private Property Get IFileSpec_Path() As String
  IFileSpec_Path = This.Path
End Property
Private Property Let IFileSpec_Path(ByVal Value As String)
  This.Path = AddProperEndingToPath(Value)
End Property

Private Property Get IFileSpec_Self() As IFileSpec
  Set IFileSpec_Self = Me
End Property

'===============================================================================

Private Function AddProperEndingToPath(ByVal Path As String) As String
  If Right$(Path, 1) <> "\" Then AddProperEndingToPath = Path & "\" _
  Else AddProperEndingToPath = Path
End Function

Private Function GetFileExt(FileName As String)
  GetFileExt = Right$(FileName, _
    Switch _
      (InStr(FileName, ".") = 0, _
        0, _
      InStr(FileName, ".") > 0, _
        Len(FileName) - InStrRev(FileName, ".")))
End Function

Private Function AddDotIfNotEmpty(ByVal Ext As String) As String
  If Ext = "" Then AddDotIfNotEmpty = "" Else AddDotIfNotEmpty = "." & Ext
End Function

Private Function GetFileName(ByVal FullFileName As String)
  GetFileName = Right$(FullFileName, Len(FullFileName) - InStrRev(FullFileName, "\"))
End Function

Private Function GetFileNameWithoutExt(ByVal FileName$) As String
  GetFileNameWithoutExt = Left$(FileName, _
    Switch _
      (InStr(FileName, ".") = 0, _
        Len(FileName), _
      InStr(FileName, ".") > 0, _
        InStrRev(FileName, ".") - 1))
End Function

Private Function GetFilePath(ByVal FullFileName As String)
  GetFilePath = Left$(FullFileName, InStrRev(FullFileName, "\"))
End Function
