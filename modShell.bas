Attribute VB_Name = "modShell"
Option Explicit

Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMINIMIZED As Long = 2
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOW As Long = 5
Public Const SW_SHOWDEFAULT As Long = 10

Public Const CREATE_NO_WINDOW As Long = &H8000000
Global Const INFINITE As Long = -1&

Private LastProcessID As Long

Private Const DIRSEP As String = "\"

Global Const NORMAL_PRIORITY_CLASS  As Long = &H20&

Enum enSW
  enSW_HIDE = 0
  enSW_NORMAL = 1
  enSW_MAXIMIZE = 3
  enSW_MINIMIZE = 6
End Enum

Type STARTUPINFO
  Cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function RunCmdToOutput(ByVal Cmd As String, Optional ByRef ErrStr As String = "", Optional ByVal AsAdmin As Boolean = False) As String
On Error GoTo RunError
  Dim A As String, B As String, C As String
  Dim tLen As Long, Iter As Long
  A = TempFile
  B = TempFile
  If Not AsAdmin Then
    ShellAndWait "cmd /c " & Cmd & " 1> " & A & " 2> " & B, enSW_HIDE
  Else
    C = TempFile(, , ".bat")
    WriteFile C, Cmd & " 1> " & A & " 2> " & B, True
    RunFileAsAdmin C, , enSW_HIDE
  End If
  
  Iter = 0
  Const MaxIter As Long = 10
  Do While True
    tLen = FileLen(A)
    Sleep 800
    If Iter > MaxIter Or FileLen(A) = tLen Then Exit Do
    Iter = Iter + 1
  Loop
  RunCmdToOutput = ReadEntireFileAndDelete(A)
  If Iter > MaxIter Then RunCmdToOutput = RunCmdToOutput & vbCrLf2 & "<<< OUTPUT TRUNCATED >>>"
  ErrStr = ReadEntireFileAndDelete(B)
  DeleteFileIfExists C
  Exit Function
  
RunError:
  RunCmdToOutput = ""
  ErrStr = "ShellOut.RunCmdToOutput: Command Execution Error - [" & Err.Number & "] " & Err.Description
End Function

' to allow for Shell.
' This routine shells out to another application and waits for it to exit.
Public Sub ShellAndWait(ByVal AppToRun As String, Optional ByVal SW As enSW = enSW_NORMAL)
  Dim NameOfProc As PROCESS_INFORMATION
  Dim NameStart As STARTUPINFO
  Dim RC As Long
    
On Error GoTo ErrorRoutineErr
  NameStart.Cb = Len(NameStart)
  If SW = enSW_HIDE Then
    RC = CreateProcessA(0&, AppToRun, 0&, 0&, CLng(SW), CREATE_NO_WINDOW, 0&, 0&, NameStart, NameOfProc)
  Else
    RC = CreateProcessA(0&, AppToRun, 0&, 0&, CLng(SW), NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
  End If
  LastProcessID = NameOfProc.dwProcessId
  RC = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
  RC = CloseHandle(NameOfProc.hProcess)
    
ErrorRoutineResume:
  Exit Sub
ErrorRoutineErr:
  MsgBox "AppShell.Form1.ShellAndWait: " & Err & Error
  Resume Next
End Sub

Public Function TempFile(Optional ByVal UseFolder As String = "", Optional ByVal UsePrefix As String = "tmp_", Optional ByVal Extension As String = ".tmp", Optional ByVal TestWrite As Boolean = True) As String
  Dim FN As String, Res As String
  If UseFolder <> "" And Not DirExists(UseFolder) Then UseFolder = ""
  If UseFolder = "" Then UseFolder = App.Path & DIRSEP
  If Right(UseFolder, 1) <> DIRSEP Then UseFolder = UseFolder & DIRSEP
  FN = Replace(UsePrefix & CDbl(Now) & "_" & App.ThreadID & "_" & Random(999999), ".", "_")
  Do While FileExists(UseFolder & FN & ".tmp")
    FN = FN & Chr(Random(25) + Asc("a"))
  Loop
  TempFile = UseFolder & FN & Extension
  
  If TestWrite Then
On Error GoTo TestWriteFailed
    WriteFile TempFile, "TEST", True, True
On Error GoTo TestReadFailed
    Res = ReadFile(TempFile)
    If Res <> "TEST" Then MsgBox "Test write to temp file " & TempFile & " failed." & vbCrLf & "Result (Len=" & Len(Res) & "):" & vbCrLf & Res, vbCritical
On Error GoTo TestClearFailed
    Kill TempFile
  End If
  Exit Function
  
TestWriteFailed:
  MsgBox "Failed to write temp file " & TempFile & "." & vbCrLf & Err.Description, vbCritical
  Exit Function
TestReadFailed:
  MsgBox "Failed to read temp file " & TempFile & "." & vbCrLf & Err.Description, vbCritical
  Exit Function
TestClearFailed:
  If Err.Number = 53 Then
    Err.Clear
    Resume Next
  End If
  
'BFH20160627
' Jerry wanted this commented out.  Absolutely horrible idea.
'  If IsDevelopment Then
  MsgBox "Failed to clear temp file " & TempFile & "." & vbCrLf & Err.Description, vbCritical
'  End If
  Exit Function
End Function

Public Sub RunShellExecuteAdmin(ByVal App As String, Optional ByVal nHwnd As Long = 0, Optional ByVal WindowState As Long = SW_SHOWNORMAL)
  If nHwnd = 0 Then nHwnd = GetDesktopWindow()
  LastProcessID = ShellExecute(nHwnd, "runas", App, vbNullString, vbNullString, WindowState)
'  ShellExecute nHwnd, "runas", App, Command & " /admin", vbNullString, SW_SHOWNORMAL
End Sub

Public Function RunFileAsAdmin(ByVal App As String, Optional ByVal nHwnd As Long = 0, Optional ByVal WindowState As Long = SW_SHOWNORMAL) As Boolean
'  If Not IsWinXP Then
  RunShellExecuteAdmin App, nHwnd, WindowState
'  Else
'    ShellOut App
'  End If
  RunFileAsAdmin = True
End Function


