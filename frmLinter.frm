VERSION 5.00
Begin VB.Form frmLinter 
   Caption         =   "Lint Project"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration:"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtResults 
         Height          =   4215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   6375
      End
      Begin VB.TextBox txtVBPFile 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   4215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "D&one"
         Height          =   495
         Left            =   3398
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdLint 
         Caption         =   "L&int"
         Default         =   -1  'True
         Height          =   495
         Left            =   1958
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblSrc 
         Alignment       =   1  'Right Justify
         Caption         =   "Project File:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblFile 
         Alignment       =   1  'Right Justify
         Caption         =   "Single File:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmLinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  txtVBPFile = modConfig.vbpFile
  txtFile = ""
End Sub

Private Sub cmdClose_Click()
  Close
End Sub

Private Sub cmdLint_Click()
  Dim File As String, Results As String
  
  fraConfig.Enabled = False
  If txtFile = "" Then
    Results = modQuickLint.Lint
  Else
    File = txtFile.Text
    If InStr(File, "\") = 0 Then File = Left(txtVBPFile, InStrRev(txtVBPFile, "\")) & File
    Results = modQuickLint.Lint(File)
  End If
  fraConfig.Enabled = True
  
  txtResults = IIf(Results = "", "Done.", Results)
End Sub

