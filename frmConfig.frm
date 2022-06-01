VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Config - VB6 To C#"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration:"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtAssemblyName 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   4215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Ca&ncel"
         Height          =   495
         Left            =   3720
         TabIndex        =   7
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   495
         Left            =   5160
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtVBPFile 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label lblAssemblyName 
         Alignment       =   1  'Right Justify
         Caption         =   "Assembly Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "Output Folder:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblSrc 
         Alignment       =   1  'Right Justify
         Caption         =   "Project File:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Config form


Private Sub Form_Load()
  modConfig.Hush = True
  With Me
    .txtVBPFile = modConfig.vbpFile
    .txtOutput = modConfig.OutputFolder
    .txtAssemblyName = modConfig.AssemblyName
  End With
  modConfig.Hush = False
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  modINI.INIWrite INISection_Settings, INIKey_VBPFile, txtVBPFile, INIFile
  modINI.INIWrite INISection_Settings, INIKey_OutputFolder, txtOutput, INIFile
  modINI.INIWrite INISection_Settings, INIKey_AssemblyName, txtAssemblyName, INIFile
  modConfig.LoadSettings True
  Unload Me
End Sub

Private Sub fraConfig_DblClick()
  If MsgBox("Reset to default?", vbOKCancel, "Config Reset") = vbCancel Then Exit Sub
  txtVBPFile = App.Path & "\prj.vbp"
  txtOutput = App.Path & "\quick"
  txtAssemblyName = "VB2CS"
End Sub

Private Sub txtOutput_Validate(ByRef Cancel As Boolean)
  If Dir(txtOutput, vbDirectory) = "" Then
    MsgBox "Output folder does not exist.  Please create to prevent errors."
  End If
End Sub

Private Sub txtVBPFile_Validate(ByRef Cancel As Boolean)
  If Dir(txtVBPFile) = "" Then
    MsgBox "Project file does not exist.  Please give a valid project to prevent errors."
  End If
End Sub

Private Sub txtAssemblyName_Validate(ByRef Cancel As Boolean)
  If txtAssemblyName = "" Then
    MsgBox "Please enter something for an assembly name."
  End If
End Sub

Private Sub txtVBPFile_GotFocus(): txtVBPFile.SelStart = 0: txtVBPFile.SelLength = Len(txtVBPFile): End Sub
Private Sub txtOutput_GotFocus(): txtOutput.SelStart = 0: txtOutput.SelLength = Len(txtOutput): End Sub
Private Sub txtAssemblyName_GotFocus(): txtAssemblyName.SelStart = 0: txtAssemblyName.SelLength = Len(txtAssemblyName): End Sub


