Attribute VB_Name = "modGit"
Option Explicit

Public Const Status As String = "status "
Public Const ST As String = "status "
Public Const Commit As String = "commit -m "
Public Const Push As String = "push "
Public Const Pull As String = "pull "
Public Const Branch As String = "branch "
Public Const BR As String = "branch "
Public Const Stash As String = "stash"
Public Const CheckOut As String = "checkout "

Private Function Gitfolder() As String
  Gitfolder = App.Path & "\"
End Function
Public Function GitCmd(ByVal C As String, Optional ByVal NoOutput As Boolean = False, Optional ByVal HideCommand As Boolean = False) As String
  Dim ErrSt As String
  PushDir Gitfolder
  If Not HideCommand Then GitOut "$ " & C
  GitCmd = RunCmdToOutput(C, ErrSt)
  PopDir
  If Not NoOutput Then GitOut GitCmd
  If ErrSt <> "" Then GitOut "ERR: " & ErrSt
End Function
Private Function GitOut(ByVal Msg As String) As Boolean
  Msg = Trim(Msg)
  Do While (Left(Msg, 1) = vbCr Or Left(Msg, 1) = vbLf): Msg = Mid(Msg, 2): Loop
  If Len(Msg) > 0 Then Debug.Print Msg
End Function

Public Function Git(ByVal C As String) As Boolean
  If LCase(Left(C, 4)) <> "git " Then C = "git " & C
  GitCmd C
  Git = True
End Function

Public Function GitConf(Optional ByVal vName As String, Optional ByVal vEMail As String, Optional ByVal Clear As Boolean = False)
  If Not IsIDE Then Exit Function
  
  GitCmd "git config --unset user.name", True, True
  GitCmd "git config --unset user.email", True, True
  If Clear Then
    GitCmd "git config --unset --global user.name", True
    GitCmd "git config --unset --global user.email", True
  ElseIf vName = "" Or vEMail = "" Then
    Debug.Print "user.name=" & Trim(Replace(Replace(GitCmd("git config --global user.name", True, True), vbCr, ""), vbLf, ""))
    Debug.Print "user.email=" & Trim(Replace(Replace(GitCmd("git config --global user.email", True, True), vbCr, ""), vbLf, ""))
  Else
'    GitCmd "git config --unset --global user.name", True
'    GitCmd "git config --unset --global user.email", True
    GitCmd "git config --global user.name " & vName
    GitCmd "git config --global user.email " & vEMail
  End If
End Function

Public Function GitPull(Optional ByVal withReset As Boolean = True) As Boolean
  If Not IsIDE Then Exit Function
'  If withReset Then GitReset
  If withReset Then
    GitCmd "git stash"
    GitCmd "git checkout master"
  End If
  
  GitCmd "git pull -r"
  If MsgBox("Restarting IDE in 5s...", vbOKCancel) = vbCancel Then Exit Function
'  RestartIDE
  GitPull = True
End Function

Public Function GitStatus() As String
  If Not IsIDE Then Exit Function
  GitStatus = GitCmd("git status")
End Function

Public Function GitVersion() As String
  If Not IsIDE Then Exit Function
  GitVersion = GitCmd("git --version")
End Function

Public Function HasGit() As Boolean
  If Not IsIDE Then Exit Function
  HasGit = GitVersion <> ""
End Function

Public Function GitReset(Optional ByVal Hard As Boolean = False, Optional ByVal toMaster As Boolean = False) As Boolean
  If Not IsIDE Then Exit Function
  If Not Hard Then
    GitCmd "git checkout -- ."
    If toMaster Then GitCmd "git checkout master -f"
  Else
    If toMaster Then GitCmd "git checkout master -f"
    GitCmd "git reset --hard"
    GitCmd "git pull -r --force"
  End If
'  RestartIDE
  GitReset = True
End Function

Public Function GitPush(ByVal Committer As String, ByVal CommitMessage As String) As Boolean
  If Not IsIDE Then Exit Function
  
  GitCmd "git add ."
  GitCmd "git status"
  
'  If MsgBox("Continue with Commit?", vbOKCancel + vbQuestion + vbDefaultButton1, "git push", , , 10) = vbCancel Then
'    GitCmd "git stash clear"
'    GitCmd "Stash cleared."
'    Exit Function
'  End If

  GitCmd "git commit -m """ & CommitMessage & """"
  GitCmd "git pull -r"

'  If MsgBox("Continue with Push?", vbOKCancel + vbQuestion + vbDefaultButton1, "git push", , , 10) = vbCancel Then
'    GitCmd "git stash clear"
'    GitCmd "Stash cleared."
'    Exit Function
'  End If

  GitCmd "git push", True
  GitCmd "git status"
  
  GitOut "GitPush Complete."
  
'  If MsgBox("Clear Credentials?", vbYesNo + vbExclamation + vbDefaultButton2, "Done?", , , 10) = vbYes Then
'    GitProgress True
'    GitConf Clear:=True
'    GitProgress
'  End If
'
  GitPush = True
End Function

Public Function GitLog(Optional ByVal CharLimit As Long = 3000)
  Dim Res As String
  If Not IsIDE Then Exit Function
  Res = GitCmd("git log", True)
  Res = Left(Res, CharLimit)
  Debug.Print Res
End Function

Public Function GitCommits() As Boolean
  GitCmd "git log --pretty=format:""%h - %an, %ar : %s"" -10"
  
  GitCommits = True
End Function

Public Function GitRemoteBranches() As Boolean
  GitCmd "git branch --remote --list"
  GitRemoteBranches = True
End Function
