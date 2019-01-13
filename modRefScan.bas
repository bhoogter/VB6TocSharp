Attribute VB_Name = "modRefScan"
Option Explicit

Private OutRes As String
Private cFuncRef_Name As String, cFuncRef_Value As String
Private Funcs As Collection

Private Function RefList(Optional ByVal KillRef As Boolean = False) As String
On Error Resume Next
  RefList = App.Path & "\refs.txt"
  If KillRef Then Kill RefList
End Function

Public Function ScanRefs() As Long
  Dim L
On Error Resume Next
  OutRes = ""
  ScanRefs = 0
  For Each L In Split(VBPModules(vbpFile), vbCrLf)
    If L = "" Then GoTo SkipItem
    ScanRefs = ScanRefs + ScanRefsFile(FilePath(vbpFile) & L)
SkipItem:
  Next
  RefList KillRef:=True
  WriteFile RefList, OutRes
  OutRes = ""
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
      
      F = M & ":" & G & ":" & F
      OutRes = OutRes & vbCrLf & F
      ScanRefsFile = ScanRefsFile + 1
    End If
NextLine:
  Next
End Function

Private Sub InitFuncs()
  Dim S As String, L
  If Dir(RefList) = "" Then ScanRefs
  If Not (Funcs Is Nothing) Then Exit Sub
  S = ReadEntireFile(RefList)
  Set Funcs = New Collection
On Error Resume Next
  For Each L In Split(S, vbCrLf)
    Funcs.Add L, SplitWord(L, 2, ":")
  Next
End Sub

Public Function FuncRef(ByVal FName As String) As String

  
'  Static S As String
'  If S = "" Then S = ReadEntireFile(RefList)
  
  If FName = cFuncRef_Name Then
    FuncRef = cFuncRef_Value
    Exit Function
  End If
  
'  FuncRef = RegExNMatch(S, ".*:" & FName & ":.*")
  InitFuncs
On Error Resume Next
  FuncRef = Funcs(FName)
  
  cFuncRef_Name = FName
  cFuncRef_Value = FuncRef
End Function

Public Function IsFuncRef(ByVal FName As String) As Boolean
  IsFuncRef = FuncRef(FName) <> ""
End Function

Public Function FuncRefModule(ByVal FName As String) As String
  FuncRefModule = nextBy(FuncRef(FName), ":")
End Function

Public Function FuncRefDecl(ByVal FName As String) As String
  FuncRefDecl = nextBy(FuncRef(FName), ":", 3)
End Function

Public Function FuncRefDeclTyp(ByVal FName As String) As String
  FuncRefDeclTyp = SplitWord(FuncRefDecl(FName), 1)
End Function

Public Function FuncRefDeclRet(ByVal FName As String) As String
  FuncRefDeclRet = FuncRefDecl(FName)
  FuncRefDeclRet = Trim(Mid(FuncRefDeclRet, InStrRev(FuncRefDeclRet, " ")))
  If Right(FuncRefDeclRet, 1) = ")" And Right(FuncRefDeclRet, 2) <> "()" Then FuncRefDeclRet = ""
End Function

Public Function FuncRefDeclArgs(ByVal FName As String) As String
  FuncRefDeclArgs = FuncRefDecl(FName)
  FuncRefDeclArgs = Mid(FuncRefDeclArgs, InStr(FuncRefDeclArgs, "(") + 1)
  FuncRefDeclArgs = Left(FuncRefDeclArgs, InStrRev(FuncRefDeclArgs, ")") - 1)
  FuncRefDeclArgs = Trim(FuncRefDeclArgs)
End Function

Public Function FuncRefDeclArgN(ByVal FName As String, ByVal N As Long) As String
  Dim F As String
  F = FuncRefDeclArgs(FName)
  FuncRefDeclArgN = nextBy(F, ", ", N)
End Function

Public Function FuncRefDeclArgCnt(ByVal FName As String) As Long
  Dim F As String, K As String
  F = FuncRefDeclArgs(FName)
  FuncRefDeclArgCnt = 0
  Do
    K = nextBy(F, ", ", FuncRefDeclArgCnt + 1)
    If K = "" Then Exit Function
    FuncRefDeclArgCnt = FuncRefDeclArgCnt + 1
  Loop While True
End Function

Public Function FuncRefArgType(ByVal FName, ByVal N As Long) As String
  FuncRefArgType = FuncRefDeclArgN(FName, N)
  If FuncRefArgType = "" Then Exit Function
  FuncRefArgType = SplitWord(FuncRefArgType, 2, " As ")
End Function

Public Function FuncRefArgByRef(ByVal FName, ByVal N As Long) As Boolean
  FuncRefArgByRef = Not IsInStr(FuncRefDeclArgN(FName, N), "ByVal ")
End Function

Public Function FuncRefArgOptional(ByVal FName, ByVal N As Long) As Boolean
  FuncRefArgOptional = IsInStr(FuncRefDeclArgN(FName, N), "Optional ")
End Function

Public Function FuncRefArgDefault(ByVal FName, ByVal N As Long) As String
  Dim aTyp As String
  If Not FuncRefArgOptional(FName, N) Then Exit Function
  FuncRefArgDefault = SplitWord(FuncRefDeclArgN(FName, N), 2, " = ", True, True)
  If FuncRefArgDefault = "" Then FuncRefArgDefault = ConvertDefaultDefault(FuncRefArgType(FName, N))
End Function


