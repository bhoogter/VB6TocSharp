Attribute VB_Name = "modDirStack"
Option Explicit

Private DirStack As New Collection

Public Function PushDir(ByVal NewDir As String, Optional ByVal doSet As Boolean = True) As String
'::::PushDir
':::SUMMARY
':Basic Directory Stack - Push cur dir to stack and CD to parameter.
':::DESCRIPTION
':1. Push Current Dir to stack
':2. CD to new folder.
':::PAREMETERS
': - sNewDir - String - Directory to CD into.
': - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
':::RETURNS
':Returns current directory.
':::SEE ALSO
': PopDir, PeekDir
  Dim N As Long
  
On Error Resume Next
  If DirStack Is Nothing Then
    Set DirStack = New Collection
    DirStack.Add 0, "n"
  End If
  
  N = Val(DirStack.Item("n")) + 1
  DirStack.Remove "n"
  DirStack.Add N, "n"
  DirStack.Add CurDir, "_" & N
  
  If doSet Then ChDir NewDir
  
  PushDir = CurDir
End Function

Public Function PopDir(Optional ByVal doSet As Boolean = True) As String
'::::PopDir
':::SUMMARY
':Remove to dir from stack.  Error Safe.  Generally to change current directory.
':::DESCRIPTION
':1. Pop Dir from stack.
':2. CD to dir.
':::PAREMETERS
': - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
':::RETURNS
':Returns directory popped.
':::SEE ALSO
': PopDir, PeekDir
  Dim N As Long, V As String

On Error Resume Next
  If DirStack Is Nothing Then Exit Function
  
  N = Val(DirStack.Item("n"))
  PopDir = DirStack.Item("_" & N)
  
  If N > 1 Then
    N = N - 1
    DirStack.Remove "n"
    DirStack.Add N, "n"
  Else
    Set DirStack = Nothing
  End If
  
  If doSet Then ChDir PopDir
End Function

Public Function PeekDir(Optional ByVal doSet As Boolean = True) As String
'::::PeekDir
':::SUMMARY
':Return directory on top of stack without removing it.  Generally to change current directory.
':::DESCRIPTION
':1. Push Current Dir to stack
':2. CD to new folder.
':::PAREMETERS
': - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
':::RETURNS
':Returns top stack item (without removing it from stack).
':::SEE ALSO
': PopDir, PeekDir
  Dim N As Long, V As String

On Error Resume Next
  If DirStack Is Nothing Then Exit Function
  
  N = Val(DirStack.Item("n"))
  PeekDir = DirStack.Item("_" & N)
  
  If doSet Then ChDir PeekDir
End Function


