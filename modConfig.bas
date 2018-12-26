Attribute VB_Name = "modConfig"
Option Explicit

Public Const SpIndent As Long = 2

Public Const PackagePrefix As String = ""
Private Const def_vbpFile As String = "C:\wincds\wincds\wincds.vbp"


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


