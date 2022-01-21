VERSION 5.00
Begin VB.Form frm 
   Caption         =   "VB6 -> .NET"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.OptionButton optVersion 
         Caption         =   "v2"
         Height          =   255
         Index           =   1
         Left            =   2648
         TabIndex        =   17
         Top             =   1080
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optVersion 
         Caption         =   "v1"
         Height          =   255
         Index           =   0
         Left            =   1928
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdSupport 
         Caption         =   "SUPPORT"
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "SCAN"
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "     Single File   ----->"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton cmdLint 
         Caption         =   "L&int"
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Conf&ig"
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtStats 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   1695
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton cmdClasses 
         Caption         =   "Classes"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdModules 
         Caption         =   "Modules"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ALL"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdForms 
         Caption         =   "Forms"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   495
         Left            =   3240
         TabIndex        =   11
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtSrc 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "C:\WinCDS\WinCDS\WinCDS.vbp"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblPrg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Shape shpPrgBack 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   2040
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Shape shpPrg 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   2040
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblSrc 
         Alignment       =   1  'Right Justify
         Caption         =   "Project File:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Main form

Public pMax As Long

Public Property Get ConverterVersion() As String
  Dim I As Long
  For I = optVersion.LBound To optVersion.UBound
    If optVersion(I) Then ConverterVersion = optVersion(I).Caption: Exit Function
  Next
  ConverterVersion = CONVERTER_VERSION_1
End Property

Private Sub cmdAll_Click()
  If Not ConfigValid Then Exit Sub
  IsWorking
  ConvertProject txtSrc, ConverterVersion
  IsWorking True
  MsgBox "Complete"
End Sub

Private Sub cmdClasses_Click()
  If Not ConfigValid Then Exit Sub
  IsWorking
  ConvertFileList FilePath(txtSrc), VBPClasses(txtSrc), vbCrLf, ConverterVersion
  IsWorking True
End Sub

Private Sub cmdConfig_Click()
  frmConfig.Show 1
  modConfig.LoadSettings
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdFile_Click()
  Dim Success As Boolean
  If txtFile = "" Then
    MsgBox "Enter a file in the box.", vbExclamation, "No File Entered"
    Exit Sub
  End If
  If Not ConfigValid Then Exit Sub
  IsWorking
  Success = ConvertFile(txtFile, False, ConverterVersion)
  IsWorking True
  If Success Then MsgBox "Converted " & txtFile & "."
End Sub

Private Sub cmdForms_Click()
  If Not ConfigValid Then Exit Sub
  IsWorking
  ConvertFileList FilePath(txtSrc), VBPForms(txtSrc), vbCrLf, ConverterVersion
  IsWorking True
End Sub

Private Sub cmdModules_Click()
  If Not ConfigValid Then Exit Sub
  IsWorking
  ConvertFileList FilePath(txtSrc), VBPModules(txtSrc), vbCrLf, ConverterVersion
  IsWorking True
End Sub

Private Function ConfigValid() As Boolean
  modConfig.LoadSettings

  If Dir(modConfig.vbpFile) = "" Then
    MsgBox "Project file not found.  Perhaps do config first?", vbExclamation, "File Not Found"
    Exit Function
  End If
  If Dir(modConfig.OutputFolder, vbDirectory) = "" Then
    MsgBox "Ouptut Folder not found.  Perhaps do config first?", vbExclamation, "Directory Not Found"
    Exit Function
  End If
  If modConfig.AssemblyName = "" Then
    MsgBox "Assembly name not set.  Perhaps do config first?", vbExclamation, "Setting Not Found"
    Exit Function
  End If
  ConfigValid = True
End Function

Private Sub IsWorking(Optional ByVal Done As Boolean = False)
  txtFile.Enabled = Done
  cmdConfig.Enabled = Done
  cmdLint.Enabled = Done
  cmdFile.Enabled = Done
  cmdAll.Enabled = Done
  cmdClasses.Enabled = Done
  cmdExit.Enabled = Done
  cmdForms.Enabled = Done
  cmdModules.Enabled = Done
  txtSrc.Enabled = Done
  cmdScan.Enabled = Done
  cmdSupport.Enabled = Done
  MousePointer = IIf(Done, vbDefault, vbHourglass)
End Sub

Public Function Prg(Optional ByVal Val As Long = -1, Optional ByVal Max As Long = -1, Optional ByVal Cap As String = "#") As String
On Error Resume Next
  If Max >= 0 Then pMax = Max
  lblPrg = IIf(Prg = "#", "", Cap)
  shpPrg.Width = Val / pMax * 2415
  shpPrg.Visible = Val >= 0
  lblPrg.Visible = shpPrg.Visible
End Function

Private Sub cmdLint_Click()
  If Not ConfigValid Then Exit Sub
  frmLinter.Show vbModal
End Sub

Private Sub cmdScan_Click()
  If Not ConfigValid Then Exit Sub
  IsWorking False
  ScanRefs
  IsWorking True
End Sub

Private Sub cmdSupport_Click()
  If Not ConfigValid Then Exit Sub
  If MsgBox("Generate Project files?", vbYesNo) = vbYes Then CreateProjectFile vbpFile
  If MsgBox("Generate Support files?", vbYesNo) = vbYes Then CreateProjectSupportFiles
End Sub

Private Sub Form_Load()
  modConfig.Hush = True
  modConfig.LoadSettings
  modConfig.Hush = False
  txtSrc = vbpFile
End Sub

