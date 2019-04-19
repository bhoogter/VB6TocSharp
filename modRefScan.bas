Attribute VB_Name = "modRefScan"
Option Explicit

Private OutRes As String
Private cFuncRef_Name As String, cFuncRef_Value As String
Private cEnuRef_Name As String, cEnumRef_Value As String
Private Funcs As Collection, LocalFuncs As Collection

Private Function RefList(Optional ByVal KillRef As Boolean = False) As String
On Error Resume Next
  RefList = App.Path & "\refs.txt"
  If KillRef Then Kill RefList
End Function

Public Function ScanRefs() As Long
  Dim L, T As String
On Error Resume Next
  OutRes = ""
  ScanRefs = 0
  For Each L In Split(VBPModules(vbpFile), vbCrLf)
    If L = "" Then GoTo SkipMod
    ScanRefs = ScanRefs + ScanRefsFile(FilePath(vbpFile) & L)
SkipMod:
  Next
  
  For Each L In Split(VBPForms(vbpFile), vbCrLf)
    L = Replace(L, ".frm", "")
    If L = "" Then GoTo SkipForm
    T = vbCrLf & L & ":" & L & ":Form:"
    OutRes = OutRes & T
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim S As String, J As Long, Preamble As String, ControlRefs As String
    S = ReadEntireFile(vbpPath & L & ".frm")
    J = CodeSectionLoc(S)
    Preamble = Left(S, J - 1)
    ControlRefs = FormControls(L, Preamble, False)
    OutRes = OutRes & ControlRefs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ScanRefs = ScanRefs + 1
SkipForm:
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
  Dim CurrEnum As String
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
      
    If tLMatch(L, "Function ") Or tLMatch(L, "Public Function ") Or _
       tLMatch(L, "Sub ") Or tLMatch(L, "Public Sub ") Or _
       False Then
      
      F = Trim(L)
      If Left(F, 7) = "Public " Then F = Mid(F, 8)
      F = Trim(nextBy(F, ":"))
      
      G = F
      If tLMatch(G, "Function ") Then G = Mid(G, 10)
      If tLMatch(G, "Sub ") Then G = Mid(G, 5)
      G = nextBy(G, "(")
      
      F = M & ":" & G & ":Function:" & F
      OutRes = OutRes & vbCrLf & F
      ScanRefsFile = ScanRefsFile + 1
    ElseIf tLMatch(L, "Declare ") Or tLMatch(L, "Public Decalre ") Then
      L = LTrim(L)
      If LMatch(L, "Public ") Then L = Mid(L, 8)
      If LMatch(L, "Declare ") Then L = Mid(L, 9)
      G = SplitWord(L)
      
    ElseIf tLMatch(L, "Const ") Or tLMatch(L, "Public Const ") Or tLMatch(L, "Global Const ") Then
      L = LTrim(L)
      If LMatch(L, "Public ") Then L = Mid(L, 8)
      If LMatch(L, "Global ") Then L = Mid(L, 8)
      If LMatch(L, "Const ") Then L = Mid(L, 7)
      G = SplitWord(L)
    ElseIf tLMatch(L, "Enum ") Or tLMatch(L, "Public Enum ") Then
      L = LTrim(L)
      If LMatch(L, "Public ") Then L = Mid(L, 8)
      If LMatch(L, "Enum ") Then L = Mid(L, 5)
      CurrEnum = Trim(L)
    ElseIf tLMatch(L, "End Enum") Then
      CurrEnum = ""
    ElseIf CurrEnum <> "" Then
      G = SplitWord(L)
      F = M & ":" & G & ":Enum:" & CurrEnum & "." & G
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
  InitLocalFuncs
End Sub

Public Sub InitLocalFuncs(Optional ByVal S As String)
On Error Resume Next
  Dim L As Variant
  Set LocalFuncs = New Collection
  For Each L In Split(S, vbCrLf)
    LocalFuncs.Add L, SplitWord(L, 2, ":")
  Next
End Sub

Public Function FuncRef(ByVal FName As String) As String
  If FName = cFuncRef_Name Then
    FuncRef = cFuncRef_Value
    Exit Function
  End If
  
  InitFuncs
On Error Resume Next
  FuncRef = Funcs(FName)
  If FuncRef = "" Then FuncRef = LocalFuncs(FName)
  
  cFuncRef_Name = FName
  cFuncRef_Value = FuncRef
End Function

Public Function FuncRefModule(ByVal FName As String) As String
  FuncRefModule = nextBy(FuncRef(FName), ":")
End Function

Public Function FuncRefEntity(ByVal FName As String) As String
  FuncRefEntity = nextBy(FuncRef(FName), ":", 3)
End Function

Public Function FuncRefDecl(ByVal FName As String) As String
  FuncRefDecl = nextBy(FuncRef(FName), ":", 4)
End Function

Public Function IsFuncRef(ByVal FName As String) As Boolean
  IsFuncRef = FuncRef(FName) <> "" And FuncRefEntity(FName) = "Function"
End Function

Public Function IsEnumRef(ByVal FName As String) As Boolean
  IsEnumRef = FuncRef(FName) <> "" And FuncRefEntity(FName) = "Enum"
End Function

Public Function IsFormRef(ByVal FName As String) As Boolean
  Dim T As String
  T = SplitWord(FName, 1, ".")
  IsFormRef = FuncRef(T) <> "" And FuncRefEntity(T) = "Form"
End Function

Public Function IsControlRef(ByVal Src As String, Optional ByVal FormName As String) As Boolean
  Dim Tok As String, Tok2 As String
  Dim FTok As String, TTok As String
  Tok = RegExNMatch(Src, patToken)
  Tok2 = RegExNMatch(Src, patToken, 1)
  TTok = Tok & "." & Tok2
  FTok = FormName & "." & Tok
'If IsInStr(Src, "SetFocus") Then Stop
  If FuncRef(TTok) <> "" And FuncRefEntity(TTok) = "Control" Or FuncRef(FTok) <> "" And FuncRefEntity(FTok) = "Control" Then
    IsControlRef = True
  End If
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
On Error Resume Next
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

Public Function FuncRefArgType(ByVal FName As String, ByVal N As Long) As String
  FuncRefArgType = FuncRefDeclArgN(FName, N)
  If FuncRefArgType = "" Then Exit Function
  FuncRefArgType = SplitWord(FuncRefArgType, 2, " As ")
End Function

Public Function FuncRefArgByRef(ByVal FName As String, ByVal N As Long) As Boolean
  FuncRefArgByRef = Not IsInStr(FuncRefDeclArgN(FName, N), "ByVal ")
End Function

Public Function FuncRefArgOptional(ByVal FName As String, ByVal N As Long) As Boolean
  FuncRefArgOptional = IsInStr(FuncRefDeclArgN(FName, N), "Optional ")
End Function

Public Function FuncRefArgDefault(ByVal FName As String, ByVal N As Long) As String
  Dim aTyp As String
  If Not FuncRefArgOptional(FName, N) Then Exit Function
  FuncRefArgDefault = SplitWord(FuncRefDeclArgN(FName, N), 2, " = ", True, True)
  If FuncRefArgDefault = "" Then FuncRefArgDefault = ConvertDefaultDefault(FuncRefArgType(FName, N))
End Function

Public Function EnumRefRepl(ByVal EName As String) As String
  EnumRefRepl = FuncRefDecl(EName)
End Function

Public Function FormRefRepl(ByVal FName As String) As String
  Dim T As String, U As String
  T = SplitWord(FName, 1, ".")
  U = FuncRefModule(T) & ".instance"
  FormRefRepl = Replace(FName, T, U)
End Function

Public Function FormControlRepl(ByVal Src As String, Optional ByVal FormName As String) As String
  Dim Tok As String, Tok2 As String, Tok3 As String
  Dim F As String, V As String
  Tok = RegExNMatch(Src, patToken)
  Tok2 = RegExNMatch(Src, patToken, 1)
  Tok3 = RegExNMatch(Src, patToken, 2)
  
'If IsInStr(Tok, "BillOSale") Then Stop
'If IsInStr(Src, "SetFocus") Then Stop
  
  If Not IsFormRef(Tok) Then
    F = Tok
    V = ConvertControlProperty(F, Tok2, FuncRefDecl(FormName & "." & Tok))
    If Tok2 <> "" Then
      FormControlRepl = Replace(Src, Tok2, V)
    Else
      FormControlRepl = Src & "." & V
    End If
  Else
    F = Tok & "." & Tok2
    V = ConvertControlProperty(F, Tok3, FuncRefDecl(Tok & "." & Tok2))
    If Tok3 <> "" Then
      FormControlRepl = Replace(Src, Tok3, V)
    Else
      FormControlRepl = Src & "." & V
    End If
  End If
End Function
