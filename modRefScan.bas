Attribute VB_Name = "modRefScan"
Option Explicit

Private Function RefList(Optional ByVal KillRef As Boolean = False) As String
On Error Resume Next
  RefList = App.Path & "\refs.txt"
  If KillRef Then Kill RefList
End Function

Public Function ScanRefs() As Long
  Dim L
On Error Resume Next
  RefList KillRef:=True
  ScanRefs = 0
  For Each L In Split(VBPModules(vbpFile) & vbCrLf & VBPClasses(vbpFile) & vbCrLf & VBPForms(vbpFile), vbCrLf)
    If L = "" Then GoTo SkipItem
    ScanRefs = ScanRefs + ScanRefsFile(FilePath(vbpFile) & L)
SkipItem:
  Next
End Function

Private Function ScanRefsFile(ByVal FN As String) As Long
  Dim M As String
  Dim S As String, L As String, LL
  Dim F As String, G As String
  Dim Cont As Boolean, DoCont As Boolean
  M = FileBaseName(FN)
  S = ReadEntireFile(FN)
  ScanRefsFile = 0
  For Each LL In Split(S, vbCrLf)
    DoCont = Right(LL, 1) = "_"
    If Not Cont And Not DoCont Then
      L = Trim(LL)
      Cont = False
    ElseIf Cont And Not DoCont Then
      L = L & Trim(LL)
      Cont = False
    ElseIf Not Cont And DoCont Then
      L = Trim(Left(LL, Len(LL) - 2))
      Cont = True
      GoTo NextLine
    ElseIf Cont And DoCont Then
      L = L & Trim(Left(LL, Len(LL) - 2))
      Cont = True
      GoTo NextLine
    End If
      
    If tLMatch(L, "Function ") Or tLMatch(L, "Public Function") Or _
       tLMatch(L, "Sub ") Or tLMatch(L, "Public Sub") Or _
       False Then
      
      F = Trim(L)
      If Left(F, 7) = "Public " Then F = Mid(F, 8)
      F = Trim(nextBy(F, ":"))
      
      G = F
      If Left(G, 9) = "Function " Then G = Mid(G, 10)
      If Left(G, 4) = "Sub " Then G = Mid(G, 5)
      G = nextBy(G, "(")
      
      WriteFile RefList, M & ":" & G & ":" & F
      ScanRefsFile = ScanRefsFile + 1
    End If
NextLine:
  Next
End Function

Public Function FuncRef(ByVal FName As String) As String
'  Dim S As String, L
  Static S As String
  If S = "" Then S = ReadEntireFile(RefList)
  FuncRef = RegExNMatch(S, "^.*:" & FName & ":.*$")
End Function
