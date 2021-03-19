Attribute VB_Name = "modConfig"
Option Explicit

Public Const SpIndent As Long = 2
Public Const DefaultDataType As String = "dynamic"

Public Const PackagePrefix As String = ""
Private Const def_vbpFile As String = "C:\WinCDS.NET\cnv\prj.vbp"
Private Const def_outputFolder As String = "C:\WinCDS.NET\cnv\converted\"
Private Const def_AssemblyName As String = "VB2CS"

Private mVBPFile As String
Private mOutputFolder As String
Private mAssemblyName As String

Private Loaded As Boolean
Public Hush As Boolean

Public Const INISection_Settings   As String = "Settings"
Public Const INIKey_VBPFile As String = "VBPFile"
Public Const INIKey_OutputFolder As String = "OutputFolder"
Public Const INIKey_AssemblyName As String = "AssemblyName"

Public Function INIFile() As String
  INIFile = App.Path & "\VB6toCS.INI"
End Function

Public Sub LoadSettings(Optional ByVal Force As Boolean)
  If Loaded And Not Force Then Exit Sub
  Loaded = True
  mVBPFile = modINI.INIRead(INISection_Settings, INIKey_VBPFile, INIFile)
  mOutputFolder = modINI.INIRead(INISection_Settings, INIKey_OutputFolder, INIFile)
  mAssemblyName = modINI.INIRead(INISection_Settings, INIKey_AssemblyName, INIFile)
End Sub

Public Function OutputFolder(Optional ByVal F As String) As String
  LoadSettings
  If mOutputFolder = "" Then mOutputFolder = def_outputFolder
  OutputFolder = mOutputFolder
  If Right(OutputFolder, 1) <> "\" Then OutputFolder = OutputFolder & "\"
  OutputFolder = OutputFolder & OutputSubFolder(F)
  If Dir(OutputFolder, vbDirectory) = "" Then
On Error GoTo CantMakeOutputFolder
    MkDir OutputFolder
  End If
  Exit Function
CantMakeOutputFolder:
  If Not Hush Then MsgBox "Failed creating folder.  Perhaps create it yourself?" & vbCrLf & OutputFolder
End Function

Public Function AssemblyName() As String
  LoadSettings
  If mAssemblyName = "" Then mAssemblyName = def_AssemblyName
  AssemblyName = mAssemblyName
End Function

Public Function OutputSubFolder(ByVal F As String) As String
  LoadSettings
  Select Case FileExt(F)
    Case ".bas": OutputSubFolder = "Modules\"
    Case ".cls": OutputSubFolder = "Classes\"
    Case ".frm": OutputSubFolder = "Forms\"
    Case Else:   OutputSubFolder = ""
  End Select
End Function

Public Property Get vbpFile() As String
  LoadSettings
  If mVBPFile = "" Then mVBPFile = def_vbpFile
  vbpFile = mVBPFile
End Property

Public Property Get vbpPath() As String
  vbpPath = FilePath(vbpFile)
End Property


