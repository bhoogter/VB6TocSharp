Attribute VB_Name = "modConvert"
Option Explicit

Const SpIndent As Long = 2
Const WithMark = "_WithVar"
Private EOLComment As String
Dim WithLevel As Long, MaxWithLevel As Long

Public Function ConvertProject(ByVal vbpFile As String)
  CreateProjectSupportFiles
  ConvertFileList VBPModules(vbpFile) & vbCrLf & VBPClasses(vbpFile) & vbCrLf & VBPForms(vbpFile) & vbCrLf & VBPUserControls(vbpFile)
  MsgBox "Complete."
End Function

Public Function ConvertFileList(ByVal Path As String, ByVal List As String, Optional ByVal Sep As String = vbCrLf) As Boolean
  Dim L, K As String
  frm.Prg 0, StrCnt(K, Sep) + 1
  For Each L In Split(List, Sep)
    If L = "" Then GoTo NextItem
    
    Select Case LCase(Right(L, 4))
      Case ".bas": ConvertModule Path & L
      Case ".cls": ConvertClass Path & L
      Case ".frm": ConvertForm Path & L
'      Case ".ctl": ConvertModule Path & L
      Case Else: MsgBox "UNKNOWN VB TYPE: " & L
    End Select
    ConvertModule Path & L
NextItem:
    N = N + 1
    frm.Prg N
  Next
End Function

Public Function ConvertForm(ByVal frmFile As String)
  Dim S As String, J As Long, Preamble As String, Code As String, Globals As String, Functions As String
  If Not FileExists(frmFile) Then
    MsgBox "File not found in ConvertForm: " & frmFile
    Exit Function
  End If
  S = ReadEntireFile(frmFile)
  J = CodeSectionLoc(S)
  Preamble = Left(S, J - 1)
  Code = Mid(S, J)
  
  X = ConvertFormUI(Preamble)
  F = DebugFolder & Replace(Mid(basFile, InStrRev(basFile, "\") + 1), ".bas", ".xaml")
  WriteFile F, X, True
  
  J = CodeSectionGlobalEndLoc(Code)
  Globals = ConvertGlobals(Left(Code, J))
  Functions = ConvertCodeSegment(Mid(Code, J))
  
  X = Globals & vbCrLf & vbCrLf & Functions
  F = DebugFolder & Replace(Mid(basFile, InStrRev(basFile, "\") + 1), ".bas", ".xaml.cs")
  WriteFile F, X, True
End Function


Public Function ConvertModule(ByVal basFile As String)
  Dim S As String, J As Long, Code As String, Globals As String, Functions As String
  Dim F As String, X As String
  If Not FileExists(basFile) Then
    MsgBox "File not found in ConvertModule: " & basFile
    Exit Function
  End If
  S = ReadEntireFile(basFile)
  Code = Mid(S, CodeSectionLoc(S))
  
  J = CodeSectionGlobalEndLoc(Code)
  Globals = ConvertGlobals(Left(Code, J))
  Functions = ConvertCodeSegment(Mid(Code, J))
  
  X = Globals & vbCrLf & vbCrLf & Functions
  F = DebugFolder & Replace(Mid(basFile, InStrRev(basFile, "\") + 1), ".bas", ".cs")
  WriteFile F, X, True
End Function



Public Function ConvertClass(ByVal clsFile As String)
  Dim S As String
  If Not FileExists(clsFile) Then
    MsgBox "File not found in ConvertForm: " & clsFile
    Exit Function
  End If
  S = ReadEntireFile(clsFile)
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
    
    FinishSplitIf = False
    If Left(LTrim(L), 3) = "If " And Right(RTrim(L), 5) <> " Then" Then
      FinishSplitIf = True
      F = nextBy(L, " Then ") & " Then"
      R = R & N & F
      L = Mid(L, Len(F) + 2)
    End If
    
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
        L = Mid(L, Len(F) + 2)

      Loop While True
    End If
    
    If FinishSplitIf Then R = R & N & "End If"
NextLine:
  Next
  
  SanitizeCode = R
End Function

Public Function CreateProjectSupportFiles() As Boolean
  Dim S As String, F As String
  S = ApplicationXAML()
  F = OutputFolder & "application.xaml"
  WriteFile F, S, True
End Function

Public Function ApplicationXAML() As String
  Dim R As String, M As String, M As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "<Application x:Class=""Application"" "
  R = R & N & "xmlns = ""http://schemas.microsoft.com/winfx/2006/xaml/presentation"" "
  R = R & N & "xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"" "
  R = R & N & "xmlns:local=""clr-namespace:WpfApp1"" "
  R = R & N & "StartupUri=""MainWindow.xaml""> "
  R = R & N & "  <Application.Resources>"
  R = R & N & "  </Application.Resources>"
  R = R & N & "</Application>"

  ApplicationXAML = R
End Function

Public Function ConvertCodeSegment(ByVal S As String) As String
  ConvertCodeSegment = ConvertSub(SanitizeCode(S))
End Function

Public Function ConvertDeclare(ByVal S As String, ByVal Ind As Long, Optional ByVal isGlobal As Boolean) As String
  Dim Sp, L
  Dim pName As String, pType As String
  Dim Res As String
  Res = ""
  
  If tLeft(S, 4) = "Dim " Then S = Mid(Trim(S), 5)
  
  Sp = Split(S, ",")
  For Each L In Sp
    L = Trim(L)
    pName = SplitWord(L, 1)
    If SplitWord(L, 2) = "As" Then
      pType = SplitWord(L, 3)
    Else
      pType = "Variant"
    End If
    
    Res = Res & sSpace(Ind) & ConvertDataType(pType) & " " & pName & ";" & vbCrLf
  Next
  
  ConvertDeclare = Res
End Function

Public Function ConvertAPIDef(ByVal S As String) As String
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
  Dim Res As String
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): isPrivate = True
  If tLeft(S, 5) = "Enum " Then S = tMid(S, 6)
  eName = RegExNMatch(S, patToken, 0)
  S = nlTrims(tMid(S, Len(eName) + 1))
  
  Res = "enum " & eName & "{" & vbCrLf
  eName = RegExNMatch(S, patToken, 0)
  S = Trim(tMid(S, Len(eName) + 1))
  Res = Res & "}"
  
End Function

Public Function ConvertType(ByVal S As String)

End Function

Public Function DeComment(ByVal Str As String) As String
  EOLComment = nextBy(Str, "'", 2)
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
  Dim isByRef As Boolean
  Dim Res As String
  Dim pName As String, pType As String, pDef As String
  
  If tLeft(S, 9) = "Optional " Then isOptional = True: S = Mid(S, 10)
  isByRef = True
  If tLeft(S, 6) = "ByRef " Then isByRef = True: S = Mid(S, 7)
  If tLeft(S, 6) = "ByVal " Then isByRef = False: S = Mid(S, 7)
  pName = SplitWord(S, 1)
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
  If isByRef Then Res = Res & "ref "
  Res = Res & ConvertDataType(pType) & " "
  Res = Res & pName & " "
  If isOptional Then
    Res = Res & "= " & pDef
  End If
  
  ConvertParameter = Trim(Res)
End Function

Public Function ConvertDefaultDefault(ByVal dType As String) As String
  Select Case dType
    Case "Long":      ConvertDefaultDefault = 0
    Case "Date":      ConvertDefaultDefault = "#1/1/2001#"
    Case "String":    ConvertDefaultDefault = """"""
    Case Else:        ConvertDefaultDefault = "null"
  End Select
End Function

Public Function ConvertDataType(ByVal S As String) As String
  Select Case S
    Case "String": ConvertDataType = "string"
    Case "Long":  ConvertDataType = "int"
    Case "Double": ConvertDataType = "double"
    Case "Variant": ConvertDataType = "object"
    Case Else: ConvertDataType = "object"
  End Select
End Function

Public Function ConvertPrototype(ByVal S As String, Optional ByRef returnVariable As String) As String
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
  S = Trim(S)
  FirstToken = RegExNMatch(S, "[a-zA-Z_][a-zA-Z_0-9.]*", 0)
  FirstWord = SplitWord(S, 1)
  If S = FirstWord Then ConvertValue = S: Exit Function
  If S = FirstToken Then ConvertValue = S & "(): exit function"
  
  If FirstToken = FirstWord And Not isOperator(SplitWord(S, 2)) Then ' Sub without parenthesis
    ConvertValue = FirstWord & "(" & SplitWord(S, 2, , , True) & ")"
  Else
    ConvertValue = S
  End If
  
  ConvertValue = Replace(ConvertValue, " & ", " + ")
  ConvertValue = Replace(ConvertValue, ":=", ": ")
  ConvertValue = Replace(ConvertValue, "=", "==")
  ConvertValue = Replace(ConvertValue, "<>", "!=")
  ConvertValue = Replace(ConvertValue, " Not ", " !")
  ConvertValue = Replace(ConvertValue, " Or ", " || ")
  ConvertValue = Replace(ConvertValue, " And ", " && ")
  ConvertValue = Replace(ConvertValue, " Mod ", " % ")
  ConvertValue = Replace(ConvertValue, " &H", "0x")

  If WithLevel > 0 Then
    ConvertValue = Trim(Replace(" " & ConvertValue, " .", " " & WithMark & WithLevel & "."))
  End If
End Function

Public Function ConvertCodeLine(ByVal S As String) As String
  Dim T As Long
  If Trim(S) = "" Then ConvertCodeLine = "": Exit Function
  
  If S Like "* = *" Then
    T = InStr(S, "=")
    ConvertCodeLine = Trim(Left(S, T - 1)) & " = " & ConvertValue(Trim(Mid(S, T + 1)))
  Else
    ConvertCodeLine = ConvertValue(S)
  End If
  
  ConvertCodeLine = ConvertCodeLine & ";"
End Function

Public Function ConvertGlobals(ByVal Str As String) As String
  Dim Res As String
  Dim S, L, O As String
  Dim Ind As Long
  Dim Building As String
  Dim inCase As Long
  Dim returnVariable As String
  
  Res = ""
  Building = ""
  Str = Replace(Str, vbLf, "")
  S = Split(Str, vbCr)
  Ind = 0
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
    ElseIf RegExTest(L, "(Public |Private |)Declare ") Then
      O = ConvertAPIDef(L)
    ElseIf RegExTest(L, "(Public |Private |)Const ") Then
      O = ConvertConstant(L, True)
    ElseIf tLeft(L, 8) = "Private " Or tLeft(L, 7) = "Public " Or tLeft(L, 4) = "Dim " Then
      O = ConvertDeclare(L, 0, True)
    ElseIf RegExTest(L, "(Public |Private |)Enum ") Then
      Building = L
    ElseIf RegExTest(L, "(Public |Private |)Type ") Then
      Building = L
    End If
      
    O = ReComment(O)
    Res = Res & ReComment(O) & IIf(O = "", "", vbCrLf)
  Next
  
  ConvertGlobals = Res
End Function

Public Function ConvertSub(ByVal Str As String)
  Dim Res As String
  Dim S, L, O As String
  Dim Ind As Long
  Dim inCase As Long
  Dim returnVariable As String
  
  Res = ""
  Str = Replace(Str, vbLf, "")
  S = Split(Str, vbCr)
  Ind = 0
  For Each L In S
    L = DeComment(L)
    O = ""

    If L Like "*Sub *" Or L Like "*Function *" Then
      O = sSpace(Ind) & ConvertPrototype(L, returnVariable)
      Ind = Ind + SpIndent
    ElseIf L Like "End Sub" Or L Like "End Function" Then
      If returnVariable <> "" Then
        O = O & sSpace(Ind) & "return " & returnVariable & ";" & vbCrLf
      End If
      Ind = Ind - SpIndent
      O = O & sSpace(Ind) & "}"
    ElseIf tLeft(L, 13) = "Exit Function" Or tLeft(L, 8) = "Exit Sub" Then
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
      O = sSpace(Ind) & ConvertConstant(L)
    ElseIf tLeft(L, 3) = "If " Then  ' Code sanitization prevents all single-line ifs.
      O = sSpace(Ind) & "if (" & ConvertValue(Mid(Trim(L), 4, Len(Trim(L)) - 8)) & ") {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 4) = "Else" Then
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
      If inCase > 0 Then O = O & sSpace(Ind) & "break;" & vbCrLf: Ind = Ind - SpIndent: inCase = inCase - 1
      O = O & sSpace(Ind) & "case " & ConvertValue(tMid(L, 6)) & ":"
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
      O = O & sSpace(Ind) & "foreach(" & SplitWord(L, 1, " In ") & " in " & SplitWord(L, 2, " In ") & ") {"
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
    ElseIf tLeft(L, 8) = "Exit For" Or tLeft(L, 7) = "Exit Do" Or tLeft(L, 10) = "Exit While" Then
      O = O & sSpace(Ind) & "break;"
    ElseIf tLeft(L, 4) = "Next" Then
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
      O = sSpace(Ind) & ConvertCodeLine(L)
    End If
    O = ReComment(O)
    Res = Res & ReComment(O) & IIf(O = "", "", vbCrLf)
  Next
  
  ConvertSub = Res
End Function

Function ConvertFormUI(ByVal S As String)
  Dim Sp, I As Long, L As String
  Dim R As String, N As String, M As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "<Window x:Class=""MainWindow"" "
  R = R & M & "xmlns = ""http://schemas.microsoft.com/winfx/2006/xaml/presentation"" "
  R = R & M & "xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"" "
  R = R & M & "xmlns:d=""http://schemas.microsoft.com/expression/blend/2008"" "
  R = R & M & "xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" "
  R = R & M & "xmlns:local=""clr-namespace:WpfApp1"" "
  R = R & M & "mc:Ignorable=""d"" "
  R = R & M & "Title=""MainWindow"" "
  R = R & M & "Height=""450"" "
  R = R & M & "Width=""800"">"
  R = R & N & "    <Grid>"
  
  Sp = Split(S, vbCrLf)
  For I = LBound(Sp) To UBound(Sp)
    L = Sp(I)
        
  Next
  
  R = R & N & "    </Grid>"
  R = R & N & "</Window>"

  ConvertFormUI = R
End Function
