Attribute VB_Name = "modConfig"
Option Explicit

Public Const SpIndent As Long = 2
Public Const DefaultDataType As String = "dynamic"

Public Const PackagePrefix As String = ""
Private Const def_vbpFile As String = "C:\wincds\wincds\wincds.vbp"
Private Const def_outputFolder As String = "C:\WinCDS.NET\out\WinCDS.NET\"

Public Function OutputFolder(Optional ByVal F As String) As String
  OutputFolder = def_outputFolder & OutputSubFolder(F)
End Function

Public Function OutputSubFolder(ByVal F As String) As String
  Select Case FileExt(F)
    Case ".bas": OutputSubFolder = "Modules\"
    Case ".cls": OutputSubFolder = "Classes\"
    Case ".frm": OutputSubFolder = "Forms\"
    Case Else:   OutputSubFolder = ""
  End Select
End Function

Public Property Get vbpFile() As String
  If Forms.Count > 0 Then
    vbpFile = frm.txtSrc
  Else
    vbpFile = def_vbpFile
  End If
End Property

Public Property Get vbpPath() As String
  vbpPath = FilePath(vbpFile)
End Property


