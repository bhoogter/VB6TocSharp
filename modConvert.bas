Attribute VB_Name = "modConvert"
Option Explicit


Const WithMark = "_WithVar"
Private EOLComment As String

Dim WithLevel As Long, MaxWithLevel As Long

Public Function ConvertProject(ByVal vbpFile As String)
  CreateProjectFile vbpFile
  CreateProjectSupportFiles
  ConvertFileList FilePath(vbpFile), VBPModules(vbpFile) & vbCrLf & VBPClasses(vbpFile) & vbCrLf & VBPForms(vbpFile) & vbCrLf & VBPUserControls(vbpFile)
  MsgBox "Complete."
End Function

Public Function ConvertFileList(ByVal Path As String, ByVal List As String, Optional ByVal Sep As String = vbCrLf) As Boolean
  Dim L, V As Long, N As Long
  V = StrCnt(List, Sep) + 1
  Prg 0, V, N & "/" & V & "..."
  For Each L In Split(List, Sep)
    N = N + 1
    If L = "" Then GoTo NextItem
    
    If L = "modFunctionList.bas" Then GoTo NextItem
    
    ConvertFile Path & L
    
NextItem:
    Prg N, , N & "/" & V & ": " & L
    DoEvents
  Next
  Prg
End Function

Public Function ConvertFile(ByVal someFile As String, Optional ByVal UIOnly As Boolean = False) As Boolean
  Select Case LCase(FileExt(someFile))
    Case ".bas": ConvertFile = ConvertModule(someFile)
    Case ".cls": ConvertFile = ConvertClass(someFile)
    Case ".frm": ConvertFile = ConvertForm(someFile, UIOnly)
'      Case ".ctl": ConvertModule  someFile
    Case Else: MsgBox "UNKNOWN VB TYPE: " & someFile
  End Select
End Function

Public Function ConvertForm(ByVal frmFile As String, Optional ByVal UIOnly As Boolean = False) As Boolean
  Dim S As String, J As Long, Preamble As String, Code As String, Globals As String, Functions As String
  Dim X As String, fName As String
  Dim F As String
  If Not FileExists(frmFile) Then
    MsgBox "File not found in ConvertForm: " & frmFile
    Exit Function
  End If
  S = ReadEntireFile(frmFile)
  fName = ModuleName(S)
  
  J = CodeSectionLoc(S)
  Preamble = Left(S, J - 1)
  Code = Mid(S, J)
  
  X = ConvertFormUi(Preamble)
  F = fName & ".xaml"
  WriteOut F, X, frmFile
  If UIOnly Then Exit Function
  
  J = CodeSectionGlobalEndLoc(Code)
  Globals = ConvertGlobals(Left(Code, J))
  Functions = ConvertCodeSegment(Mid(Code, J))
  
  X = ""
  X = X & UsingEverything(fName) & vbCrLf
  X = X & vbCrLf
  X = X & "public class " & fName & " {" & vbCrLf
  X = X & "  public static " & fName & " DefaultInstance;" & vbCrLf
  X = X & Globals & vbCrLf & vbCrLf & Functions
  X = X & vbCrLf & "}"
  
  X = deWS(X)
  
  F = fName & ".xaml.cs"
  WriteOut F, X, frmFile
End Function


Public Function ConvertModule(ByVal basFile As String)
  Dim S As String, J As Long, Code As String, Globals As String, Functions As String
  Dim F As String, X As String, fName As String
  If Not FileExists(basFile) Then
    MsgBox "File not found in ConvertModule: " & basFile
    Exit Function
  End If
  S = ReadEntireFile(basFile)
  fName = ModuleName(S)
  Code = Mid(S, CodeSectionLoc(S))
  
  J = CodeSectionGlobalEndLoc(Code)
  Globals = ConvertGlobals(Left(Code, J - 1))
  Functions = ConvertCodeSegment(Mid(Code, J), True)
  
  X = ""
  X = X & UsingEverything(fName) & vbCrLf
  X = X & vbCrLf
  X = X & "static class " & fName & " {" & vbCrLf
  X = X & nlTrim(Globals & vbCrLf & vbCrLf & Functions)
  X = X & vbCrLf & "}"
  F = fName & ".cs"
  WriteOut F, X, basFile
End Function



Public Function ConvertClass(ByVal clsFile As String)
  Dim S As String, J As Long, Code As String, Globals As String, Functions As String
  Dim F As String, X As String, fName As String
  Dim cName As String
  If Not FileExists(clsFile) Then
    MsgBox "File not found in ConvertModule: " & clsFile
    Exit Function
  End If
  S = ReadEntireFile(clsFile)
  fName = ModuleName(S)
  Code = Mid(S, CodeSectionLoc(S))
  
  J = CodeSectionGlobalEndLoc(Code)
  Globals = ConvertGlobals(Left(Code, J - 1))
  Functions = ConvertCodeSegment(Mid(Code, J))
  
  X = ""
  X = X & UsingEverything(fName) & vbCrLf
  X = X & vbCrLf
  X = X & "public class " & fName & " {" & vbCrLf
  X = X & Globals & vbCrLf & vbCrLf & Functions
  X = X & vbCrLf & "}"
  
  F = fName & ".cs"
  WriteOut F, X, clsFile
End Function

Public Function SanitizeCode(ByVal Str As String)
  Const NamedParamSrc = ":="
  Const NamedParamTok = "###NAMED-PARAMETER###"
  Dim Sp, L
  Dim F As String
  Dim R As String, N As String
  Dim Building As String
  Dim FinishSplitIf As Boolean
  
  R = "": N = vbCrLf
  Sp = Split(Str, vbCrLf)
  Building = ""
  For Each L In Sp
    If Right(L, 1) = "_" Then Building = Building & Trim(Left(L, Len(L) - 1)) & " ": GoTo NextLine
    If Building <> "" Then
      L = Building & Trim(L)
      Building = ""
    End If
    
    L = DeComment(L)
'If IsInStr(L, "CustRec <> 0") Then Stop
    
    FinishSplitIf = False
    If tLeft(L, 3) = "If " And Right(RTrim(L), 5) <> " Then" Then
      FinishSplitIf = True
      F = nextBy(L, " Then ") & " Then"
      R = R & N & F
      L = Mid(L, Len(F) + 2)
      If nextBy(L, " Else ", 2) <> "" Then
        R = R & N & SanitizeCode(nextBy(L, " Else ", 1))
        R = R & N & "Else"
        L = nextBy(L, "Else ", 2)
      End If
    End If
    
    If nextBy(L, ":") <> L Then
      If RegExTest(Trim(L), "^[a-zA-Z_][a-zA-Z_0-9]*:$") Then ' Goto Label
        R = R & N & ReComment(L)
      Else
        Do
          L = Replace(L, NamedParamSrc, NamedParamTok)
          F = nextBy(L, ":")
          F = Replace(F, NamedParamTok, NamedParamSrc)
          R = R & N & ReComment(F, True)
          L = Replace(L, NamedParamTok, NamedParamSrc)
          If F = L Then Exit Do
          L = Trim(Mid(L, Len(F) + 2))
          R = R & N & SanitizeCode(L)

        Loop While False
      End If
    Else
      R = R & N & L
    End If
    
    If FinishSplitIf Then R = R & N & "End If"
NextLine:
  Next
  
  SanitizeCode = R
End Function

Public Function ConvertCodeSegment(ByVal S As String, Optional ByVal AsModule As Boolean = False) As String
  Dim P As String, N As Long
  Dim F As String, T As Long, E As Long, K As String, X As Long
  Dim Pre As String, Body As String
  Dim R As String
  
  ClearProperties
  
  S = SanitizeCode(S)
  Do
    P = "(Public |Private |)(Function |Sub |Property Get |Property Let |Property Set )" & patToken & "\("
    N = -1
    Do
      N = N + 1
      F = RegExNMatch(S, P, N)
      T = RegExNPos(S, P, N)
    Loop While Not IsInCode(S, T) And F <> ""
    If F = "" Then Exit Do
    
    If IsInStr(F, " Function ") Then K = "End Function"
    If IsInStr(F, " Sub ") Then K = "End Sub"
    If IsInStr(F, " Property ") Then K = "End Property"
    N = -1
    Do
      N = N + 1
      E = RegExNPos(Mid(S, T), K, N) + Len(K) + T
    Loop While Not IsInCode(S, E) And E <> 0
    
    If T > 1 Then Pre = nlTrim(Left(S, T - 1)) Else Pre = ""
    Do Until Mid(S, E, 1) = vbCr Or Mid(S, E, 1) = vbLf Or Mid(S, E, 1) = ""
      E = E + 1
    Loop
    Body = nlTrim(Mid(S, T, E - T))
      
    S = nlTrim(Mid(S, E + 1))
    
    R = R & CommentBlock(Pre) & ConvertSub(Body, AsModule) & vbCrLf
  Loop While True
  
  R = ReadOutProperties & vbCrLf2 & R
  
  ConvertCodeSegment = R
End Function

Public Function CommentBlock(ByVal Str As String) As String
  Dim S As String
  If nlTrim(Str) = "" Then Exit Function
  S = ""
  S = S & "/*" & vbCrLf
  S = S & Replace(Str, "*/", "* /") & vbCrLf
  S = S & "*/" & vbCrLf
  CommentBlock = S
End Function

Public Function ConvertDeclare(ByVal S As String, ByVal Ind As Long, Optional ByVal isGlobal As Boolean) As String
  Dim Sp, L, SS As String
  Dim asPrivate As Boolean
  Dim pName As String, pType As String, pWithEvents As Boolean
  Dim Res As String
  Dim ArraySpec As String, isArr As Boolean, aMax As String, aMin As String, aTodo As String
  Res = ""
  
  SS = S
  
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 4) = "Dim " Then S = Mid(Trim(S), 5): asPrivate = True
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): asPrivate = True
  
  Sp = Split(S, ",")
  For Each L In Sp
    L = Trim(L)
    If LMatch(L, "WithEvents ") Then L = Trim(tMid(L, 12)): Res = Res & "// TODO: WithEvents not supported on " & RegExNMatch(L, patToken) & vbCrLf
    pName = RegExNMatch(L, patToken)
    L = Trim(tMid(L, Len(pName) + 1))
    If isGlobal Then Res = Res & IIf(asPrivate, "private ", "public ")
    If tLeft(L, 1) = "(" Then
      isArr = True
      ArraySpec = nextBy(Mid(L, 2), ")")
      If ArraySpec = "" Then
        aMin = -1
        aMax = -1
        L = Trim(tMid(L, 3))
      Else
        L = Trim(tMid(L, Len(ArraySpec) + 3))
        aMin = 0
        aMax = SplitWord(ArraySpec)
        ArraySpec = Trim(tMid(ArraySpec, Len(aMax) + 1))
        If tLeft(ArraySpec, 3) = "To " Then
          aMin = aMax
          aMax = tMid(ArraySpec, 4)
        End If
      End If
    End If
    If SplitWord(L, 1) = "As" Then
      pType = SplitWord(L, 2)
    Else
      pType = "Variant"
    End If
    
    If Not isArr Then
      Res = Res & sSpace(Ind) & ConvertDataType(pType) & " " & pName & ";" & vbCrLf
    Else
      aTodo = IIf(aMin = 0, "", " // TODO - Specified Minimum Array Boundary Not Supported: " & SS)
      If Not IsNumeric(aMax) Then
        Res = Res & sSpace(Ind) & "List<" & ConvertDataType(pType) & "> " & pName & " = new List<" & ConvertDataType(pType) & "> (new " & ConvertDataType(pType) & "[(" & aMax & " + 1)]);  // TODO: Confirm Array Size By Token" & aTodo & vbCrLf
      ElseIf Val(aMax) = -1 Then
        Res = Res & sSpace(Ind) & "List<" & ConvertDataType(pType) & "> " & pName & " = new List<" & ConvertDataType(pType) & "> {};" & aTodo & vbCrLf
      Else
        Res = Res & sSpace(Ind) & "List<" & ConvertDataType(pType) & "> " & pName & " = new List<" & ConvertDataType(pType) & "> (new " & ConvertDataType(pType) & "[" & (Val(aMax) + 1) & "]);" & aTodo & vbCrLf
      End If
    End If
    
    SubParamDecl pName, pType, isArr, False
  Next
  
  ConvertDeclare = Res
End Function

Public Function ConvertAPIDef(ByVal S As String) As String
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'[DllImport("User32.dll")]
'public static extern int MessageBox(int h, string m, string c, int type);
  Dim isPrivate As Boolean, isSub As Boolean
  Dim aName As String
  Dim aLib As String
  Dim aAlias As String
  Dim aArgs As String
  Dim aReturn As String
  Dim tArg As String, has As Boolean
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): isPrivate = True
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 8) = "Declare " Then S = tMid(S, 9)
  If tLeft(S, 4) = "Sub " Then S = tMid(S, 5): isSub = True
  If tLeft(S, 9) = "Function " Then S = tMid(S, 10)
  aName = RegExNMatch(S, patToken)
  S = Trim(tMid(S, Len(aName) + 1))
  If tLeft(S, 4) = "Lib " Then
    S = Trim(tMid(S, 5))
    aLib = SplitWord(S, 1)
    S = Trim(tMid(S, Len(aLib) + 1))
    If Left(aLib, 1) = """" Then aLib = Mid(aLib, 2)
    If Right(aLib, 1) = """" Then aLib = Left(aLib, Len(aLib) - 1)
    If LCase(Right(aLib, 4)) <> ".dll" Then aLib = aLib & ".dll"
    aLib = LCase(aLib)
  End If
  If tLeft(S, 6) = "Alias " Then
    S = Trim(tMid(S, 7))
    aAlias = SplitWord(S, 1)
    S = Trim(tMid(S, Len(aAlias) + 1))
    If Left(aAlias, 1) = """" Then aAlias = Mid(aAlias, 2)
    If Right(aAlias, 1) = """" Then aAlias = Left(aAlias, Len(aAlias) - 1)
    End If
  If tLeft(S, 1) = "(" Then S = tMid(S, 2)
  aArgs = nextBy(S, ")")
  S = Trim(tMid(S, Len(aArgs) + 2))
  If tLeft(S, 3) = "As " Then
    S = Trim(tMid(S, 4))
    aReturn = SplitWord(S, 1)
    S = Trim(tMid(S, Len(aReturn) + 1))
  Else
    aReturn = "Variant"
  End If
  
  S = ""
  S = S & "[DllImport(""" & aLib & """)" & IIf(aAlias = "", "", ", DllEntryPoint(""" & aAlias & """)") & "] "
  S = S & IIf(isPrivate, "private ", "public ")
  S = S & "static extern "
  S = S & IIf(isSub, "void ", ConvertDataType(aReturn))
  S = S & aName
  S = S & "("
  Do
    If aArgs = "" Then Exit Do
    tArg = Trim(nextBy(aArgs, ","))
    aArgs = tMid(aArgs, Len(tArg) + 2)
    S = S & IIf(has, ", ", "") & ConvertParameter(tArg)
    has = True
  Loop While True
  S = S & ");"
  
  
  ConvertAPIDef = S
End Function

Public Function ConvertConstant(ByVal S As String, Optional ByVal isGlobal As Boolean = True) As String
  Dim cName As String, cType As String, cVal As String, isPrivate As Boolean
  If tLeft(S, 7) = "Public " Then S = Mid(Trim(S), 8)
  If tLeft(S, 8) = "Private " Then S = Mid(Trim(S), 9): isPrivate = True
  If tLeft(S, 6) = "Const " Then S = Mid(Trim(S), 7)
  cName = SplitWord(S, 1)
  S = Trim(Mid(Trim(S), Len(cName) + 1))
  If tLeft(S, 3) = "As " Then
    S = Trim(Mid(Trim(S), 3))
    cType = SplitWord(S, 1)
    S = Trim(tMid(S, Len(cType) + 1))
  Else
    cType = "Variant"
  End If
  
  If Left(S, 1) = "=" Then
    S = Trim(Mid(S, 2))
    cVal = ConvertValue(S)
  Else
    cVal = ConvertDefaultDefault(cType)
  End If
  
  ConvertConstant = IIf(isGlobal, IIf(isPrivate, "private ", "public "), "") & "const " & ConvertDataType(cType) & " " & cName & " = " & cVal & ";"
End Function

Public Function ConvertEnum(ByVal S As String)
  Dim isPrivate As Boolean, eName As String
  Dim Res As String, has As Boolean
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): isPrivate = True
  If tLeft(S, 5) = "Enum " Then S = tMid(S, 6)
  eName = RegExNMatch(S, patToken, 0)
  S = nlTrim(tMid(S, Len(eName) + 1))
  
  Res = "enum " & eName & " {"
  
  Do While tLeft(S, 8) <> "End Enum" And S <> ""
    eName = RegExNMatch(S, patToken, 0)
    Res = Res & IIf(has, ",", "") & vbCrLf & sSpace(SpIndent) & eName
    has = True

    S = nlTrim(tMid(S, Len(eName) + 1))
    If tLeft(S, 1) = "=" Then
      S = nlTrim(Mid(S, 2))
      eName = RegExNMatch(S, "[0-9]*", 0)
      Res = Res & " = " & eName
      S = nlTrim(tMid(S, Len(eName) + 1))
    End If
  Loop
  Res = Res & vbCrLf & "}"
End Function

Public Function ConvertType(ByVal S As String)
  Dim isPrivate As Boolean, eName As String, eType As String
  Dim Res As String
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): isPrivate = True
  If tLeft(S, 5) = "Type " Then S = tMid(S, 6)
  eName = RegExNMatch(S, patToken, 0)
  S = nlTrim(tMid(S, Len(eName) + 1))
  
  Res = IIf(isPrivate, "private ", "public ") & "struct " & eName & " {"
  
  Do While tLeft(S, 8) <> "End Enum" And S <> ""
    eName = RegExNMatch(S, patToken, 0)
    S = nlTrim(tMid(S, Len(eName) + 1))
    
    If tLeft(S, 3) = "As " Then
      S = nlTrim(Mid(S, 4))
      eType = RegExNMatch(S, patToken, 0)
      S = nlTrim(tMid(S, Len(eType) + 1))
    Else
      eType = "Variant"
    End If
    Res = Res & vbCrLf & " public " & ConvertDataType(eType) & " " & eName & ";"
  Loop
  Res = Res & vbCrLf & "}"
End Function

Public Function DeComment(ByVal Str As String, Optional ByVal Discard As Boolean = False) As String
  If Not Discard Then EOLComment = nextBy(Str, "'", 2)
  DeComment = RTrim(nextBy(Str, "'", 1))
End Function

Public Function ReComment(ByVal Str As String, Optional ByVal KeepVBComments As Boolean = False)
  Dim C As String
  Dim Pr As String
  Pr = IIf(KeepVBComments, "'", "//")
  If EOLComment = "" Then ReComment = Str: Exit Function
  C = Pr & EOLComment
  EOLComment = ""
  If Not IsInStr(Str, vbCrLf) Then
    ReComment = Str & IIf(Len(Str) = 0, "", " ") & C
  Else
    ReComment = Replace(Str, vbCrLf, C & vbCrLf, , 1)         ' Always leave on end of first line...
  End If
  If Left(LTrim(ReComment), 2) = Pr Then ReComment = LTrim(ReComment)
End Function

Public Function ConvertParameter(ByVal S As String) As String
  Dim isOptional As Boolean
  Dim isByRef As Boolean, asOut As Boolean
  Dim Res As String
  Dim pName As String, pType As String, pDef As String
  
  S = Trim(S)
  If tLeft(S, 9) = "Optional " Then isOptional = True: S = Mid(S, 10)
  isByRef = True
  If tLeft(S, 6) = "ByVal " Then isByRef = False: S = Mid(S, 7)
  If tLeft(S, 6) = "ByRef " Then isByRef = True: S = Mid(S, 7)
  pName = SplitWord(S, 1)
  If isByRef And SubParam(pName).AssignedBeforeUsed Then asOut = True
  S = Trim(Mid(S, Len(pName) + 1))
  If tLeft(S, 2) = "As" Then
    S = tMid(S, 4)
    pType = SplitWord(S, 1, "=")
    S = Trim(Mid(S, Len(pType) + 1))
  Else
    pType = "Variant"
  End If
  If Left(S, 1) = "=" Then
    pDef = Trim(Mid(Trim(S), 2))
    S = ""
  Else
    pDef = ConvertDefaultDefault(pType)
  End If
  
  Res = ""
  If isByRef Then Res = Res & IIf(asOut, "out ", "ref ")
  Res = Res & ConvertDataType(pType) & " "
  Res = Res & IIf(SubParam(pName).Used, pName, pName & "_UNUSED") & " "
  If isOptional Then
    Res = Res & "= " & pDef
  End If
  
  SubParamDecl pName, pType, False, True
  ConvertParameter = Trim(Res)
End Function

Public Function ConvertPrototype(ByVal S As String, Optional ByRef returnVariable As String, Optional ByVal AsModule As Boolean = False) As String
  Const retToken = "#RET#"
  Dim Res As String
  Dim fName As String, fArgs As String, retType As String, T As String
  Dim tArg As String
  Dim isSub As Boolean
  Dim hArgs As Boolean
  
  Res = ""
  returnVariable = ""
  isSub = False
  If tLeft(S, 7) = "Public " Then Res = Res & "public ": S = Mid(S, 8)
  If tLeft(S, 8) = "Private " Then Res = Res & "private ": S = Mid(S, 9)
  If AsModule Then Res = Res & "static "
  If tLeft(S, 4) = "Sub " Then Res = Res & "void ": S = Mid(S, 5): isSub = True
  If tLeft(S, 9) = "Function " Then Res = Res & retToken & " ": S = Mid(S, 10)
  
  fName = Trim(SplitWord(Trim(S), 1, "("))
  S = Mid(S, Len(fName) + 2)
  fArgs = Trim(nextBy(S, ")"))
  S = Mid(S, Len(fArgs) + 2)
  
  If Not isSub Then
    If tLeft(S, 2) = "As" Then
      retType = Trim(Mid(Trim(S), 3))
    Else
      retType = "Variant"
    End If
    Res = Replace(Res, retToken, ConvertDataType(retType))
  End If
  
  Res = Res & fName
  Res = Res & "("
  hArgs = False
  Do
    If Trim(fArgs) = "" Then Exit Do
    tArg = nextBy(fArgs, ",")
    fArgs = LTrim(Mid(fArgs, Len(tArg) + 2))
    
    Res = Res & IIf(hArgs, ", ", "") & ConvertParameter(tArg)
    hArgs = True
  Loop Until Len(fArgs) = 0
  
  Res = Res & ") {"
  If retType <> "" Then
    returnVariable = fName
    Res = Res & vbCrLf & sSpace(SpIndent) & ConvertDataType(retType) & " " & returnVariable & " = " & ConvertDefaultDefault(retType) & ";"
  End If
  
  ConvertPrototype = Trim(Res)
End Function

Public Function ConvertCondition(ByVal S As String) As String
  ConvertCondition = "(" & S & ")"
End Function

Public Function ConvertValue(ByVal S As String) As String
  Dim FirstToken As String, FirstWord As String
  Dim T As String, Complete As Boolean
  S = Trim(S)
  
'If IsInStr(S, "RS!") Then Stop
'If IsInStr(S, ".SetValueDisplay Row") Then Stop
'If IsInStr(S, "cmdSaleTotals.Move") Then Stop
'If IsInStr(S, "2830") Then Stop
  
  S = RegExReplace(S, patNotToken & patToken & "!" & patToken & patNotToken, "$1$2(""$3"")$4")
  
  S = ConvertVb6Specific(S, Complete)
  If Complete Then ConvertValue = S: Exit Function

  SubParamUsedList TokenList(S)
  
  FirstToken = RegExNMatch(S, patTokenDot, 0)
  FirstWord = SplitWord(S, 1)
  If FirstWord = "Not" Then
    S = "!" & Mid(S, 5)
    FirstWord = SplitWord(Mid(S, 2))
  End If
  If S = FirstWord Then ConvertValue = S: GoTo DoReplacements
  If S = FirstToken Then ConvertValue = S & "()": GoTo DoReplacements
  
  If FirstToken = FirstWord And Not isOperator(SplitWord(S, 2)) Then ' Sub without parenthesis
    ConvertValue = FirstWord & "(" & SplitWord(S, 2, , , True) & ")"
  Else
    ConvertValue = S
  End If
  
  
DoReplacements:
  ConvertValue = Replace(ConvertValue, " & ", " + ")
  ConvertValue = Replace(ConvertValue, ":=", ": ")
  ConvertValue = Replace(ConvertValue, " = ", " == ")
  ConvertValue = Replace(ConvertValue, "<>", "!=")
  ConvertValue = Replace(ConvertValue, " Not ", " !")
  ConvertValue = Replace(ConvertValue, "(Not ", "(!")
  ConvertValue = Replace(ConvertValue, " Or ", " || ")
  ConvertValue = Replace(ConvertValue, " And ", " && ")
  ConvertValue = Replace(ConvertValue, " Mod ", " % ")
  ConvertValue = Replace(ConvertValue, "New ", "new ")
  Do While IsInStr(ConvertValue, ", ,")
    ConvertValue = Replace(ConvertValue, ", ,", ", _,")
  Loop
  ConvertValue = Replace(ConvertValue, "(,", "(_,")

'If IsInStr(ConvertValue, "&H") And Right(ConvertValue, 1) = "&" Then Stop
'If IsInStr(ConvertValue, "1/1/2001") Then Stop

  ConvertValue = RegExReplace(ConvertValue, "([0-9])#", "$1")
  ConvertValue = RegExReplace(ConvertValue, "#([0-9]?[0-9])/([0-9]?[0-9])/([0-9][0-9][0-9][0-9])#", """$1/$2/$3""")
  
  If Left(ConvertValue, 2) = "&H" Then
    ConvertValue = "0x" & Mid(ConvertValue, 3)
    If Right(ConvertValue, 1) = "&" Then ConvertValue = Left(ConvertValue, Len(ConvertValue) - 1)
  End If
  
  ConvertValue = ConvertStrings(ConvertValue)

  If WithLevel > 0 Then
    ConvertValue = Trim(RegExReplace(ConvertValue, "([ (])(\.)" & patToken, "$1" & WithMark & WithLevel & "$2$3"))
    If Left(ConvertValue, 1) = "." Then ConvertValue = WithMark & WithLevel & ConvertValue
  End If
End Function

Public Function TestConvertStrings()
  Const Q As String = """"
  Dim S As String
  S = "a + b + " & Q & "hello" & Q & " + " & Q & ", " & Q & Q & Q & " + k + " & Q & Q & Q & Q & " & " & Q & Q & Q & Q & Q & Q
  TestConvertStrings = ConvertStrings(S)
'  Debug.Print S
'  Debug.Print TestConvertStrings
End Function

Public Function ConvertStrings(ByVal S As String)
  Dim A As Long, B As Long
  Dim T As String, U As String
  A = InStr(S, """")
  Do Until A = 0
    B = InStr(A + 1, S, """")
    If B = 0 Then Exit Do
    Do While Mid(S, B + 1, 1) = """" And B <> 0
      B = InStr(B + 2, S, """")
    Loop
    If B = 0 Then Exit Do
    
    T = Mid(S, A, B - A + 1)
    U = """" & ConvertString(Mid(T, 2, Len(T) - 2)) & """"
    
    If T <> U Then
      S = Left(S, A - 1) & U & Mid(S, B + 1)
    End If
    A = InStr(B - Len(T) + Len(U) + 1, S, """")
  Loop
  ConvertStrings = S
End Function

Public Function ConvertString(ByVal S As String)
  S = Replace(S, "\", "\\")
  S = Replace(S, """""", "\""")
  ConvertString = S
End Function

Public Function ConvertGlobals(ByVal Str As String) As String
  Dim Res As String
  Dim S, L, O As String
  Dim Ind As Long
  Dim Building As String
  Dim inCase As Long
  Dim returnVariable As String
  Dim N As Long
  
  Res = ""
  Building = ""
  Str = Replace(Str, vbLf, "")
  S = Split(Str, vbCr)
  Ind = 0
  N = 0
'  Prg 0, UBound(S) - LBound(S) + 1, "Globals..."
  For Each L In S
    L = DeComment(L)
    O = ""
    If Building <> "" Then
      Building = Building & vbCrLf & L
      If tLeft(L, 8) = "End Type" Then
        O = ConvertType(Building)
        Building = ""
      ElseIf tLeft(L, 8) = "End Enum" Then
        O = ConvertEnum(Building)
        Building = ""
      End If
    ElseIf L Like "Option *" Then
      O = "// " & L
    ElseIf RegExTest(L, "([^a-zA-Z0-9_])(Public |Private |)Declare ") Then
      O = ConvertAPIDef(L)
    ElseIf RegExTest(L, "([^a-zA-Z0-9_])(Public |Private |)Const ") Then
      O = ConvertConstant(L, True)
    ElseIf RegExTest(L, "([^a-zA-Z0-9_])(Public |Private |)Enum ") Then
      Building = L
    ElseIf RegExTest(L, "([^a-zA-Z0-9_])(Public |Private |)Type ") Then
      Building = L
    ElseIf tLeft(L, 8) = "Private " Or tLeft(L, 7) = "Public " Or tLeft(L, 4) = "Dim " Then
      O = ConvertDeclare(L, 0, True)
    End If
      
    O = ReComment(O)
    Res = Res & ReComment(O) & IIf(O = "" Or Right(O, 2) = vbCrLf, "", vbCrLf)
    N = N + 1
'    Prg N
'    If N Mod 10000 = 0 Then Stop
  Next
'  Prg
  
  ConvertGlobals = Res
End Function

Public Function ConvertCodeLine(ByVal S As String) As String
  Dim T As Long, A As String
  If Trim(S) = "" Then ConvertCodeLine = "": Exit Function
  
  If RegExTest(Trim(S), "^[a-zA-Z0-9_.()""]+ \= ") Or RegExTest(Trim(S), "^Set [a-zA-Z0-9_.]+ \= ") Then
    T = InStr(S, "=")
    A = Trim(Left(S, T - 1))
    If tLeft(A, 4) = "Set " Then A = Trim(Mid(A, 5))
    SubParamAssign A
    If Left(A, 1) = "." Then A = WithMark & WithLevel & A
    ConvertCodeLine = A & " = " & ConvertValue(Trim(Mid(S, T + 1)))
  Else
'Debug.Print S
      ConvertCodeLine = ConvertValue(S)
  End If
  
  ConvertCodeLine = ConvertCodeLine & ";"
'Debug.Print ConvertCodeLine
End Function



Public Function ConvertSub(ByVal Str As String, Optional ByVal AsModule As Boolean = False, Optional ByVal ScanFirst As VbTriState = vbUseDefault)
  Dim oStr As String
  Dim Res As String
  Dim S, L, O As String, T As String, U As String
  Dim cM As Long, cN As Long
  Dim K As Long
  Dim Ind As Long
  Dim inCase As Long
  Dim returnVariable As String
  
  Select Case ScanFirst
    Case vbUseDefault:  oStr = Str
                        ConvertSub oStr, AsModule, vbTrue
                        ConvertSub = ConvertSub(oStr, AsModule, vbFalse)
                        Exit Function
    Case vbTrue:        SubBegin
    Case vbFalse:       SubBegin True
  End Select
  

  
  Res = ""
  Str = Replace(Str, vbLf, "")
  S = Split(Str, vbCr)
  Ind = 0
  For Each L In S
    L = DeComment(L)
    O = ""

'If IsInStr(L, "1/1/2001") Then Stop
'If ScanFirst = vbFalse Then Stop
    If L Like "*Sub *" Or L Like "*Function *" Then
      O = sSpace(Ind) & ConvertPrototype(L, returnVariable, AsModule)
      Ind = Ind + SpIndent
    ElseIf L Like "*Property *" Then
      AddProperty Str
      Exit Function    ' repacked later...  not added here.
    ElseIf tLMatch(L, "End Sub") Or tLMatch(L, "End Function") Then
      If returnVariable <> "" Then
        O = O & sSpace(Ind) & "return " & returnVariable & ";" & vbCrLf
      End If
      Ind = Ind - SpIndent
      O = O & sSpace(Ind) & "}"
    ElseIf tLMatch(L, "Exit Function") Or tLMatch(L, "Exit Sub") Then
      If returnVariable <> "" Then
        O = O & sSpace(Ind) & "return " & returnVariable & ";" & vbCrLf
      Else
        O = O & "return;" & vbCrLf
      End If
    ElseIf RegExTest(Trim(L), "^[a-zA-Z_][a-zA-Z_0-9]*:$") Then ' Goto Label
      O = O & L
    ElseIf tLeft(L, 3) = "Dim" Then
      O = ConvertDeclare(L, Ind)
    ElseIf tLeft(L, 5) = "Const" Then
      O = sSpace(Ind) & ConvertConstant(L, False)
    ElseIf tLeft(L, 3) = "If " Then  ' Code sanitization prevents all single-line ifs.
'If IsInStr(L, "Development") Then Stop
      T = Mid(Trim(L), 4, Len(Trim(L)) - 8)
      O = sSpace(Ind) & "if (" & ConvertValue(T) & ") {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 7) = "ElseIf " Then
      T = tMid(L, 8)
      If Right(Trim(L), 5) = " Then" Then T = Left(T, Len(T) - 5)
      O = sSpace(Ind - SpIndent) & "} else if (" & ConvertValue(T) & ") {"
    ElseIf tLeft(L, 5) = "Else" Then
      O = sSpace(Ind - SpIndent) & "} else {"
    ElseIf tLeft(L, 6) = "End If" Then
      Ind = Ind - SpIndent
      O = sSpace(Ind) & "}"
    ElseIf tLeft(L, 12) = "Select Case " Then
      O = O & sSpace(Ind) & "switch(" & ConvertValue(tMid(L, 13)) & ") {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 10) = "End Select" Then
      If inCase > 0 Then Ind = Ind - SpIndent: inCase = inCase - 1
      Ind = Ind - SpIndent
      O = O & "}"
    ElseIf tLeft(L, 9) = "Case Else" Then
      If inCase > 0 Then O = O & sSpace(Ind) & "break;" & vbCrLf: Ind = Ind - SpIndent: inCase = inCase - 1
      O = O & sSpace(Ind) & "default:"
      inCase = inCase + 1
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 5) = "Case " Then
      T = Mid(Res, InStrRev(Res, "switch("))
      If RegExTest(T, "case [^:]+:") Then O = O & sSpace(Ind) & "break;" & vbCrLf: Ind = Ind - SpIndent: inCase = inCase - 1
      T = tMid(L, 6)
      If tLeft(T, 5) = "Like " Or tLeft(T, 3) = "Is " Or T Like "* = *" Then
        O = O & "// TODO: Cannot convert case: " & T & vbCrLf
        O = O & sSpace(Ind) & "case 0: "
      ElseIf nextBy(T, ",", 2) <> "" Then
        O = O & sSpace(Ind)
        Do
          U = nextBy(T, ", ")
          If U = "" Then Exit Do
          T = Trim(Mid(T, Len(U) + 1))
          O = O & "case " & ConvertValue(U) & ": "
        Loop While True
      ElseIf T Like "* To *" Then
        O = O & "// CONVERSION: Case was " & T & vbCrLf
        O = O & sSpace(Ind)
        cN = Val(SplitWord(T, 1, " To "))
        cM = Val(SplitWord(T, 2, " To "))
        For K = cN To cM
          O = O & "case " & K & ": "
        Next
      Else
        O = O & sSpace(Ind) & "case " & ConvertValue(T) & ":"
      End If
      inCase = inCase + 1
      Ind = Ind + SpIndent
    ElseIf Trim(L) = "Do" Then
      O = O & sSpace(Ind) & "do {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 9) = "Do While " Then
      O = O & sSpace(Ind) & "while(" & ConvertValue(tMid(L, 10)) & ") {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 9) = "Do Until " Then
      O = O & sSpace(Ind) & "while(!(" & ConvertValue(tMid(L, 10)) & ")) {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 9) = "For Each " Then
      L = tMid(L, 10)
      O = O & sSpace(Ind) & "foreach(var " & SplitWord(L, 1, " In ") & " in " & SplitWord(L, 2, " In ") & ") {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 4) = "For " Then
      Dim forKey As String, forStr As String, forEnd As String
      L = tMid(L, 5)
      forKey = SplitWord(L, 1, "=")
      L = SplitWord(L, 2, "=")
      forStr = SplitWord(L, 1, " To ")
      forEnd = SplitWord(L, 2, " To ")
      O = O & sSpace(Ind) & "for(" & forKey & "=" & forStr & "; " & forKey & "<" & forEnd & "; " & forKey + "++) {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 11) = "Loop While " Then
      Ind = Ind - SpIndent
      O = O & sSpace(Ind) & "} while(!(" & ConvertValue(tMid(L, 12)) & ");"
    ElseIf tLeft(L, 11) = "Loop Until " Then
      Ind = Ind - SpIndent
      O = O & sSpace(Ind) & "} while(!(" & ConvertValue(tMid(L, 12)) & ");"
    ElseIf tLeft(L, 5) = "Loop" Then
      Ind = Ind - SpIndent
      O = O & sSpace(Ind) & "}"
    ElseIf tLeft(L, 8) = "Exit For" Or tLeft(L, 7) = "Exit Do" Or tLeft(L, 10) = "Exit While" Then
      O = O & sSpace(Ind) & "break;"
    ElseIf tLeft(L, 5) = "Next" Then
      Ind = Ind - SpIndent
      O = sSpace(Ind) & "}"
    ElseIf tLeft(L, 5) = "With " Then
      WithLevel = WithLevel + 1
      If WithLevel > MaxWithLevel Then
        O = O & sSpace(Ind) & "object " & WithMark & WithLevel & ";" & vbCrLf
        MaxWithLevel = MaxWithLevel + 1
      End If
      O = O & sSpace(Ind) & WithMark & " = " & tMid(L, 6) & ";"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 8) = "End With" Then
      WithLevel = WithLevel - 1
      Ind = Ind - SpIndent
    ElseIf IsInStr(L, "On Error ") Or IsInStr(L, "Resume ") Then
      O = sSpace(Ind) & "// TODO (not supported): " & L
    Else
'If IsInStr(L, "ComputeAgeing dtpArrearControlDate") Then Stop
      O = sSpace(Ind) & ConvertCodeLine(L)
    End If
    O = ReComment(O)
    Res = Res & ReComment(O) & IIf(O = "", "", vbCrLf)
  Next
  
  ConvertSub = Res
End Function
