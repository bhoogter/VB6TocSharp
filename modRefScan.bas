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

Public Function FuncsCount(Optional ByVal vLocal As Boolean = False) As Long
On Error Resume Next
  If vLocal Then
    FuncsCount = LocalFuncs.Count
  Else
    FuncsCount = Funcs.Count
  End If
End Function

Public Function ScanRefs() As Long
  Dim L As Variant, T As String, LL As String
On Error Resume Next
  OutRes = ""
  ScanRefs = 0
  For Each L In Split(VBPModules(vbpFile), vbCrLf)
    If L = "" Then GoTo SkipMod
    LL = Replace(L, ".bas", "")
    OutRes = OutRes & vbCrLf & LL & ":" & LL & ":Module:"
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
  Dim S As String, L As String, LL As Variant
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
    ElseIf tLMatch(L, "Private Function ") Or _
       tLMatch(L, "Private Sub ") Or _
       False Then
      
      F = Trim(L)
      F = Trim(nextBy(F, ":"))
      
      G = F
      If tLMatch(G, "Private Function ") Then G = Mid(G, 17)
      If tLMatch(G, "Private Sub ") Then G = Mid(G, 12)
      G = nextBy(G, "(")
      
      F = M & ":" & Trim(M) & "." & Trim(G) & ":Private Function:" & F
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

Public Function ScanRefsFileToString(ByVal FN As String) As String
  OutRes = ""
  ScanRefsFile FN
  ScanRefsFileToString = OutRes
  OutRes = ""
End Function


Private Sub InitFuncs()
  Dim S As String, L As Variant
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

Public Sub InitLocalFuncs(Optional ByVal S As String = "")
On Error Resume Next
  Dim L As Variant
  Set LocalFuncs = New Collection
  For Each L In Split(S, vbCrLf)
    LocalFuncs.Add L, SplitWord(L, 2, ":")
  Next
End Sub

Public Function FuncRef(ByVal fName As String) As String
  If fName = cFuncRef_Name Then
    FuncRef = cFuncRef_Value
    Exit Function
  End If
  
  InitFuncs
On Error Resume Next
  FuncRef = Funcs(fName)
  If FuncRef = "" Then FuncRef = LocalFuncs(fName)
  
  cFuncRef_Name = fName
  cFuncRef_Value = FuncRef
End Function

Public Function FuncRefModule(ByVal fName As String) As String
  FuncRefModule = nextBy(FuncRef(fName), ":")
End Function

Public Function FuncRefEntity(ByVal fName As String) As String
  FuncRefEntity = nextBy(FuncRef(fName), ":", 3)
End Function

Public Function FuncRefDecl(ByVal fName As String) As String
  FuncRefDecl = nextBy(FuncRef(fName), ":", 4)
End Function

Public Function IsFuncRef(ByVal fName As String) As Boolean
  IsFuncRef = FuncRef(fName) <> "" And FuncRefEntity(fName) = "Function"
End Function

Public Function IsPrivateFuncRef(ByVal Module As String, ByVal fName As String) As Boolean
  Dim TName As String
  TName = Trim(Module) & "." & Trim(fName)
  IsPrivateFuncRef = FuncRef(TName) <> "" And FuncRefEntity(TName) = "Private Function"
End Function

Public Function IsEnumRef(ByVal fName As String) As Boolean
  IsEnumRef = FuncRef(fName) <> "" And FuncRefEntity(fName) = "Enum"
End Function

Public Function IsFormRef(ByVal fName As String) As Boolean
  Dim T As String
  T = SplitWord(fName, 1, ".")
  IsFormRef = FuncRef(T) <> "" And FuncRefEntity(T) = "Form"
End Function

Public Function IsModuleRef(ByVal fName As String) As Boolean
  Dim T As String
  T = SplitWord(fName, 1, ".")
  IsModuleRef = FuncRef(T) <> "" And FuncRefEntity(T) = "Module"
End Function

Public Function IsControlRef(ByVal Src As String, Optional ByVal FormName As String = "") As Boolean
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

Public Function FormControlRefDeclType(ByVal Src As String, Optional ByVal FormName As String = "") As String
  Dim Tok As String, Tok2 As String
  Dim FTok As String, TTok As String
  Tok = RegExNMatch(Src, patToken)
  Tok2 = RegExNMatch(Src, patToken, 1)
  TTok = Tok & "." & Tok2
  FTok = FormName & "." & Tok
'If IsInStr(Src, "SetFocus") Then Stop
  If FuncRef(TTok) <> "" And FuncRefEntity(TTok) = "Control" Then
    FormControlRefDeclType = FuncRefDecl(TTok)
  ElseIf FuncRef(FTok) <> "" And FuncRefEntity(FTok) = "Control" Then
    FormControlRefDeclType = FuncRefDecl(FTok)
  End If
End Function


Public Function FuncRefDeclTyp(ByVal fName As String) As String
  FuncRefDeclTyp = SplitWord(FuncRefDecl(fName), 1)
End Function

Public Function FuncRefDeclRet(ByVal fName As String) As String
  FuncRefDeclRet = FuncRefDecl(fName)
  FuncRefDeclRet = Trim(Mid(FuncRefDeclRet, InStrRev(FuncRefDeclRet, " ")))
  If Right(FuncRefDeclRet, 1) = ")" And Right(FuncRefDeclRet, 2) <> "()" Then FuncRefDeclRet = ""
End Function

Public Function FuncRefDeclArgs(ByVal fName As String) As String
On Error Resume Next
  FuncRefDeclArgs = FuncRefDecl(fName)
  FuncRefDeclArgs = Mid(FuncRefDeclArgs, InStr(FuncRefDeclArgs, "(") + 1)
  FuncRefDeclArgs = Left(FuncRefDeclArgs, InStrRev(FuncRefDeclArgs, ")") - 1)
  FuncRefDeclArgs = Trim(FuncRefDeclArgs)
End Function

Public Function FuncRefDeclArgN(ByVal fName As String, ByVal N As Long) As String
  Dim F As String
  F = FuncRefDeclArgs(fName)
  FuncRefDeclArgN = nextBy(F, ", ", N)
End Function

Public Function FuncRefDeclArgCnt(ByVal fName As String) As Long
  Dim F As String, K As String
  F = FuncRefDeclArgs(fName)
  FuncRefDeclArgCnt = 0
  Do
    K = nextBy(F, ", ", FuncRefDeclArgCnt + 1)
    If K = "" Then Exit Function
    FuncRefDeclArgCnt = FuncRefDeclArgCnt + 1
  Loop While True
End Function

Public Function FuncRefArgType(ByVal fName As String, ByVal N As Long) As String
  FuncRefArgType = FuncRefDeclArgN(fName, N)
  If FuncRefArgType = "" Then Exit Function
  FuncRefArgType = SplitWord(FuncRefArgType, 2, " As ")
End Function

Public Function FuncRefArgByRef(ByVal fName As String, ByVal N As Long) As Boolean
  FuncRefArgByRef = Not IsInStr(FuncRefDeclArgN(fName, N), "ByVal ")
End Function

Public Function FuncRefArgOptional(ByVal fName As String, ByVal N As Long) As Boolean
  FuncRefArgOptional = IsInStr(FuncRefDeclArgN(fName, N), "Optional ")
End Function

Public Function FuncRefArgDefault(ByVal fName As String, ByVal N As Long) As String
  Dim aTyp As String
  If Not FuncRefArgOptional(fName, N) Then Exit Function
  FuncRefArgDefault = SplitWord(FuncRefDeclArgN(fName, N), 2, " = ", True, True)
  If FuncRefArgDefault = "" Then FuncRefArgDefault = ConvertDefaultDefault(FuncRefArgType(fName, N))
End Function

Public Function EnumRefRepl(ByVal EName As String) As String
  EnumRefRepl = FuncRefDecl(EName)
End Function

Public Function FormRefRepl(ByVal fName As String) As String
  Dim T As String, U As String
  T = SplitWord(fName, 1, ".")
  U = FuncRefModule(T) & ".instance"
  FormRefRepl = Replace(fName, T, U)
End Function

Public Function FormControlRepl(ByVal Src As String, Optional ByVal FormName As String = "") As String
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
