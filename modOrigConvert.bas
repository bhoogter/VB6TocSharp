Attribute VB_Name = "modOrigConvert"
Option Explicit

Const WithMark As String = "_WithVar_"

Dim WithLevel As Long, MaxWithLevel As Long
Dim WithVars As String, WithTypes As String, WithAssign As String
Dim FormName As String

Dim CurrentModule As String


Dim CurrSub As String

Public Function GetMultiLineSpace(ByVal Prv As String, ByVal Nxt As String) As String
  Dim pC As String, nC As String
  GetMultiLineSpace = " "
  pC = Right(Prv, 1)
  nC = Left(Nxt, 1)
  If nC = "(" Then GetMultiLineSpace = ""
End Function

Public Function SanitizeCode(ByVal Str As String) As String
  Const NamedParamSrc As String = ":="
  Const NamedParamTok As String = "###NAMED-PARAMETER###"
  Dim Sp() As String, L As Variant
  Dim F As String
  Dim R As String, N As String
  Dim Building As String
  Dim FinishSplitIf As Boolean
  
  R = "": N = vbCrLf
  Sp = Split(Str, vbCrLf)
  Building = ""
  

  For Each L In Sp
'If IsInStr(L, "POEDIFolder") Then Stop
'If IsInStr(L, "Set objSourceArNo = New_CDbTypeAhead") Then Stop
    If Right(L, 1) = "_" Then
      Dim C As String
      C = Trim(Left(L, Len(L) - 1))
      Building = Building & GetMultiLineSpace(Building, C) & C
      GoTo NextLine
    End If
    If Building <> "" Then
      L = Building & GetMultiLineSpace(Building, Trim(L)) & Trim(L)
      Building = ""
    End If
    
'    If IsInStr(L, "'") Then Stop
    L = DeComment(L)
    L = DeString(L)
'If IsInStr(L, "CustRec <> 0") Then Stop
    
    FinishSplitIf = False
    If tLeft(L, 3) = "If " And Right(RTrim(L), 5) <> " Then" Then
      FinishSplitIf = True
      F = nextBy(L, " Then ") & " Then"
      R = R & N & F
      L = Mid(L, Len(F) + 2)
      If nextBy(L, " Else ", 2) <> "" Then
        R = R & SanitizeCode(nextBy(L, " Else ", 1))
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
          R = R & SanitizeCode(L)

        Loop While False
      End If
    Else
      R = R & N & ReComment(L, True)
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
  
  InitDeString
'WriteFile "C:\Users\benja\Desktop\code.txt", S, True
  S = SanitizeCode(S)
'WriteFile "C:\Users\benja\Desktop\sani.txt", S, True
  Do
    P = "(Public |Private |)(Friend |)(Function |Sub |Property Get |Property Let |Property Set )" & patToken & "[ ]*\("
    N = -1
    Do
      N = N + 1
      F = RegExNMatch(S, P, N)
      T = RegExNPos(S, P, N)
    Loop While Not IsInCode(S, T) And F <> ""
    If F = "" Then Exit Do
    
    If IsInStr(F, " Function ") Then
      K = "End Function"
    ElseIf IsInStr(F, " Sub ") Then
      K = "End Sub"
    ElseIf IsInStr(F, " Property ") Then
      K = "End Property"
    End If
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
  
  R = ReadOutProperties(AsModule) & vbCrLf2 & R
  
  R = ReString(R, True)
  
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

Public Function ConvertDeclare(ByVal S As String, ByVal Ind As Long, Optional ByVal isGlobal As Boolean = False, Optional ByVal AsModule As Boolean = False) As String
  Dim Sp() As String, L As Variant, SS As String
  Dim asPrivate As Boolean
  Dim pName As String, pType As String, pWithEvents As Boolean
  Dim Res As String
  Dim ArraySpec As String, isArr As Boolean, aMax As Long, aMin As Long, aTodo As String
  Res = ""
  
  SS = S
  
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 4) = "Dim " Then S = Mid(Trim(S), 5): asPrivate = True
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): asPrivate = True
  
'  If IsInStr(S, "aMin") Then Stop
  Sp = Split(S, ",")
  For Each L In Sp
    L = Trim(L)
    If LMatch(L, "WithEvents ") Then L = Trim(tMid(L, 12)): Res = Res & "// TODO: WithEvents not supported on " & RegExNMatch(L, patToken) & vbCrLf
    pName = RegExNMatch(L, patToken)
    L = Trim(tMid(L, Len(pName) + 1))
    If isGlobal Then Res = Res & IIf(asPrivate, "private ", "public ")
    If AsModule Then Res = Res & "static "
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
        aMax = Val(SplitWord(ArraySpec))
        ArraySpec = Trim(tMid(ArraySpec, Len(aMax) + 1))
        If tLeft(ArraySpec, 3) = "To " Then
          aMin = aMax
          aMax = Val(tMid(ArraySpec, 4))
        End If
      End If
    End If
    
    Dim AsNew As Boolean
    AsNew = False
    If SplitWord(L, 1) = "As" Then
      pType = SplitWord(L, 2)
      If pType = "New" Then
        pType = SplitWord(L, 3)
        AsNew = True
      End If
    Else
      pType = "Variant"
    End If
    
    If Not isArr Then
      Res = Res & sSpace(Ind) & ConvertDataType(pType) & " " & pName
      Res = Res & " = "
      If AsNew Then
        Res = Res & "new "
        Res = Res & ConvertDataType(pType)
        Res = Res & "()"
      Else
        Res = Res & ConvertDefaultDefault(pType)
      End If
      Res = Res & ";" & vbCrLf
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
    
    SubParamDecl pName, pType, IIf(isArr, "" & aMax, ""), False, False
  Next
  
  ConvertDeclare = Res
End Function

Public Function ConvertAPIDef(ByVal S As String) As String
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'[DllImport("User32.dll")]
'public static extern int MessageBox(int h, string m, string c, int type);
  Dim isPrivate As Boolean, IsSub As Boolean
  Dim AName As String
  Dim aLib As String
  Dim aAlias As String
  Dim aArgs As String
  Dim aReturn As String
  Dim tArg As String, Has As Boolean
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): isPrivate = True
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 8) = "Declare " Then S = tMid(S, 9)
  If tLeft(S, 4) = "Sub " Then S = tMid(S, 5): IsSub = True
  If tLeft(S, 9) = "Function " Then S = tMid(S, 10)
  AName = RegExNMatch(S, patToken)
  S = Trim(tMid(S, Len(AName) + 1))
  If tLeft(S, 4) = "Lib " Then
    S = Trim(tMid(S, 5))
    aLib = SplitWord(S, 1)
    S = Trim(tMid(S, Len(aLib) + 1))
    aLib = ReString(aLib)
    If Left(aLib, 1) = """" Then aLib = Mid(aLib, 2)
    If Right(aLib, 1) = """" Then aLib = Left(aLib, Len(aLib) - 1)
    If LCase(Right(aLib, 4)) <> ".dll" Then aLib = aLib & ".dll"
    aLib = LCase(aLib)
  End If
  If tLeft(S, 6) = "Alias " Then
    S = Trim(tMid(S, 7))
    aAlias = SplitWord(S, 1)
    S = Trim(tMid(S, Len(aAlias) + 1))
    aAlias = ReString(aAlias)
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
  S = S & "[DllImport(""" & aLib & """" & IIf(aAlias = "", "", ", EntryPoint = """ & aAlias & """") & ")] "
  S = S & IIf(isPrivate, "private ", "public ")
  S = S & "static extern "
  S = S & IIf(IsSub, "void ", ConvertDataType(aReturn)) & " "
  S = S & AName
  S = S & "("
  Do
    If aArgs = "" Then Exit Do
    tArg = Trim(nextBy(aArgs, ","))
    aArgs = tMid(aArgs, Len(tArg) + 2)
    S = S & IIf(Has, ", ", "") & ConvertParameter(tArg, True)
    Has = True
  Loop While True
  S = S & ");"
  
  
  ConvertAPIDef = S
End Function

Public Function ConvertConstant(ByVal S As String, Optional ByVal isGlobal As Boolean = True) As String
  Dim cName As String, cType As String, cValue As String, isPrivate As Boolean, dataType As String
  If tLeft(S, 7) = "Public " Then S = Mid(Trim(S), 8)
  If tLeft(S, 7) = "Global " Then S = Mid(Trim(S), 8)
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
    cValue = ConvertValue(S)
  Else
    cValue = ConvertDefaultDefault(cType)
  End If
  
  dataType = ConvertDataType(cType)
  If dataType = "dynamic" Then ' c# can't handle constants of type 'dynamic' when type can be inferred.
    If LMatch(cValue, DeStringToken_Base) Then
      dataType = "string"
    ElseIf IsNumeric(cValue) Then
      If IsInStr(cValue, ".") Then dataType = "decimal" Else dataType = "int"
    End If
  End If
  
  If cType = "Date" Then
    ConvertConstant = IIf(isGlobal, IIf(isPrivate, "private ", "public "), "") & "static readonly " & dataType & " " & cName & " = " & cValue & ";"
  Else
    ConvertConstant = IIf(isGlobal, IIf(isPrivate, "private ", "public "), "") & "const " & dataType & " " & cName & " = " & cValue & ";"
  End If
End Function


Public Function ConvertEvent(ByVal S As String) As String
  Dim cName As String, cArgs As String, tArgs As String, isPrivate As Boolean
  Dim R As String, N As Long, M As String, O As String
  Dim I As Long, J As Long
  Dim A As String
  If tLeft(S, 7) = "Public " Then S = Mid(Trim(S), 8)
  If tLeft(S, 8) = "Private " Then S = Mid(Trim(S), 9): isPrivate = True
  If tLeft(S, 6) = "Event " Then S = Mid(Trim(S), 7)
  cName = RegExNMatch(S, patToken)
  cArgs = Trim(Mid(Trim(S), Len(cName) + 1))
  If Left(cArgs, 1) = "(" Then cArgs = Mid(cArgs, 2)
  If Right(cArgs, 1) = ")" Then cArgs = Left(cArgs, Len(cArgs) - 1)
  
  N = 0
  Do
    N = N + 1
    A = nextBy(cArgs, ",", N)
    If A = "" Then Exit Do
    tArgs = tArgs & IIf(N = 1, "", ", ")
    tArgs = tArgs & ConvertParameter(A, True)
  Loop While True
  
  O = vbCrLf
  M = ""
  R = ""
  R = R & M & "public delegate void " & cName & "Handler(" & tArgs & ");"
  R = R & O & "public event " & cName & "Handler event" & cName & ";"
  
  ConvertEvent = R
End Function


Public Function ConvertEnum(ByVal S As String) As String
  Dim isPrivate As Boolean, EName As String
  Dim Res As String, Has As Boolean
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): isPrivate = True
  If tLeft(S, 5) = "Enum " Then S = tMid(S, 6)
  EName = RegExNMatch(S, patToken, 0)
  S = nlTrim(tMid(S, Len(EName) + 1))
  
  Res = "public enum " & EName & " {"
  
  Do While tLeft(S, 8) <> "End Enum" And S <> ""
    EName = RegExNMatch(S, patToken, 0)
    Res = Res & IIf(Has, ",", "") & vbCrLf & sSpace(SpIndent) & EName
    Has = True

    S = nlTrim(tMid(S, Len(EName) + 1))
    If tLeft(S, 1) = "=" Then
      S = nlTrim(Mid(S, 3))
      If Left(S, 1) = "&" Then
        EName = ConvertElement(RegExNMatch(S, "&H[0-9A-F]+"))
      Else
        EName = RegExNMatch(S, "[0-9]*", 0)
      End If
      Res = Res & " = " & EName
      S = nlTrim(tMid(S, Len(EName) + 1))
    End If
  Loop
  Res = Res & vbCrLf & "}"
  
  ConvertEnum = Res
End Function

Public Function ConvertType(ByVal S As String) As String
  Dim isPrivate As Boolean, EName As String, eArr As String, eType As String
  Dim Res As String
  Dim N As String
  If tLeft(S, 7) = "Public " Then S = tMid(S, 8)
  If tLeft(S, 8) = "Private " Then S = tMid(S, 9): isPrivate = True
  If tLeft(S, 5) = "Type " Then S = tMid(S, 6)
  EName = RegExNMatch(S, patToken, 0)
  S = nlTrim(tMid(S, Len(EName) + 1))
'If IsInStr(eName, "OSVERSIONINFO") Then Stop
  
  Res = IIf(isPrivate, "private ", "public ") & "class " & EName & " {"
  
  Do While tLeft(S, 8) <> "End Type" And S <> ""
    EName = RegExNMatch(S, patToken, 0)
    S = nlTrim(tMid(S, Len(EName) + 1))
    eArr = ""
    If LMatch(S, "(") Then
      N = nextBy(Mid(S, 2), ")")
      S = nlTrim(Mid(S, Len(N) + 3))
      N = ConvertValue(N)
      eArr = "[" & N & "]"
    End If
    
    If tLeft(S, 3) = "As " Then
      S = nlTrim(Mid(S, 4))
      eType = RegExNMatch(S, patToken, 0)
      S = nlTrim(tMid(S, Len(eType) + 1))
    Else
      eType = "Variant"
    End If
    Res = Res & vbCrLf & " public " & ConvertDataType(eType) & IIf(eArr = "", "", "[]") & " " & EName
    If eArr = "" Then
      Res = Res & " = " & ConvertDefaultDefault(eType)
    Else
      Res = Res & " = new " & ConvertDataType(eType) & eArr
    End If
    Res = Res & ";"
    If tLMatch(S, "* ") Then
      S = Mid(LTrim(S), 3)
      N = RegExNMatch(S, "[0-9]+", 0)
      S = nlTrim(Mid(LTrim(S), Len(N) + 1))
      Res = Res & " //TODO: Fixed Length Strings Not Supported: * " & N
    End If

  Loop
  Res = Res & vbCrLf & "}"
  
  ConvertType = Res
End Function

Public Function ConvertParameter(ByVal S As String, Optional ByVal NeverUnused As Boolean = False) As String
  Dim IsOptional As Boolean
  Dim IsByRef As Boolean, asOut As Boolean
  Dim Res As String
  Dim pName As String, pType As String, pDef As String
  Dim TName As String
  
  S = Trim(S)
  If tLeft(S, 9) = "Optional " Then IsOptional = True: S = Mid(S, 10)
  IsByRef = True
  If tLeft(S, 6) = "ByVal " Then IsByRef = False: S = Mid(S, 7)
  If tLeft(S, 6) = "ByRef " Then IsByRef = True: S = Mid(S, 7)
  pName = SplitWord(S, 1)
  If IsByRef And SubParam(pName).AssignedBeforeUsed Then asOut = True
  S = Trim(Mid(S, Len(pName) + 1))
  If tLeft(S, 2) = "As" Then
    S = tMid(S, 4)
    pType = SplitWord(S, 1, "=")
    S = Trim(Mid(S, Len(pType) + 1))
  Else
    pType = "Variant"
  End If
  If Left(S, 1) = "=" Then
    pDef = ConvertValue(Trim(Mid(Trim(S), 2)))
    S = ""
  Else
    pDef = ConvertDefaultDefault(pType)
  End If
  
  Res = ""
  If IsByRef Then Res = Res & IIf(asOut, "out ", "ref ")
  Res = Res & ConvertDataType(pType) & " "
  If IsInStr(pName, "()") Then Res = Res & "[] ": pName = Replace(pName, "()", "")
  TName = pName
  If Not NeverUnused Then
    If Not SubParam(pName).Used And Not (SubParam(pName).Param And SubParam(pName).Assigned) Then
      TName = TName & "_UNUSED"
    End If
  End If
  Res = Res & TName
  If IsOptional And Not IsByRef Then
    Res = Res & "= " & pDef
  End If
  
  SubParamDecl pName, pType, False, True, False
  ConvertParameter = Trim(Res)
End Function

Public Function ConvertPrototype(ByVal SS As String, Optional ByRef returnVariable As String = "", Optional ByVal AsModule As Boolean = False, Optional ByRef asName As String = "") As String
  Const retToken As String = "#RET#"
  Dim Res As String
  Dim fName As String, fArgs As String, retType As String, T As String
  Dim tArg As String
  Dim IsSub As Boolean
  Dim hArgs As Boolean
  Dim S As String
  
  S = SS
  
  Res = ""
  returnVariable = ""
  IsSub = False
  If LMatch(S, "Public ") Then Res = Res & "public ": S = Mid(S, 8)
  If LMatch(S, "Private ") Then Res = Res & "private ": S = Mid(S, 9)
  If LMatch(S, "Friend ") Then S = Mid(S, 8)
  If AsModule Then Res = Res & "static "
  If LMatch(S, "Sub ") Then Res = Res & "void ": S = Mid(S, 5): IsSub = True
  If LMatch(S, "Function ") Then Res = Res & retToken & " ": S = Mid(S, 10)
  
  fName = Trim(SplitWord(Trim(S), 1, "("))
  asName = fName
  
  S = Trim(tMid(S, Len(fName) + 2))
  If Left(S, 1) = "(" Then S = Trim(tMid(S, 2))
  fArgs = Trim(nextBy(S, ")"))
  S = Mid(S, Len(fArgs) + 2)
  Do While Right(fArgs, 1) = "("
    fArgs = fArgs & ") "
    Dim tMore As String
    tMore = Trim(nextBy(S, ")"))
    fArgs = fArgs & tMore
    S = Mid(S, Len(tMore) + 2)
  Loop
  If Left(S, 1) = ")" Then S = Trim(tMid(S, 2))
  
  If Not IsSub Then
    If tLeft(S, 2) = "As" Then
      retType = Trim(Mid(Trim(S), 3))
    Else
      retType = "Variant"
    End If
    If Right(retType, 1) = ")" And Right(retType, 2) <> "()" Then retType = Left(retType, Len(retType) - 1)
    Res = Replace(Res, retToken, ConvertDataType(retType))
  End If
  
  Res = Res & fName
  Res = Res & "("
  hArgs = False
  Do
    If Trim(fArgs) = "" Then Exit Do
    tArg = nextBy(fArgs, ",")
    fArgs = LTrim(Mid(fArgs, Len(tArg) + 2))
    
    Res = Res & IIf(hArgs, ", ", "")
    If LMatch(tArg, "ParamArray") Then Res = Res & "params ": tArg = "ByVal " & Trim(Mid(tArg, 12))
    Res = Res & ConvertParameter(tArg)
    hArgs = True
  Loop Until Len(fArgs) = 0
  
  Res = Res & ") {"
  If retType <> "" Then
    returnVariable = fName
    Res = Res & vbCrLf & sSpace(SpIndent) & ConvertDataType(retType) & " " & returnVariable & " = " & ConvertDefaultDefault(retType) & ";"
    SubParamDecl returnVariable, retType, False, False, True
  End If
  
  If IsEvent(asName) Then Res = EventStub(asName) & Res
  ConvertPrototype = Trim(Res)
End Function

Public Function ConvertCondition(ByVal S As String) As String
  ConvertCondition = "(" & S & ")"
End Function

Public Function ConvertElement(ByVal S As String) As String
'Debug.Print "ConvertElement: " & S
'If IsInStr(S, "frmSetup") Then Stop
'If IsInStr(S, "chkShowBalance.Value") Then Stop
'If IsInStr(S, "optTelephone") Then Stop
  Dim FirstToken As String, FirstWord As String
  Dim T As String, Complete As Boolean
  S = Trim(S)
  If S = "" Then Exit Function
  
'If IsInStr(S, "Debug.Print") Then Stop
  If Left(Trim(S), 2) = "&H" Then
    ConvertElement = "0x" & Mid(Trim(S), 3)
    Exit Function
  End If
  
  If IsNumeric(Trim(S)) Then
    ConvertElement = Val(S)
    If IsInStr(S, ".") Then ConvertElement = ConvertElement & "m"
    Exit Function
  End If
  
  Dim vMax As Long
  Do While RegExTest(S, "#[0-9]+/[0-9]+/[0-9]+#")
    Dim dStr As String
    dStr = RegExNMatch(S, "#[0-9]+/[0-9]+/[0-9]+#", 0)
    S = Replace(S, dStr, "DateValue(""" & Mid(dStr, 2, Len(dStr) - 2) & """)")
    vMax = vMax + 1
    If vMax > 10 Then Exit Do
  Loop
  
'If IsInStr(S, "RS!") Then Stop
'If IsInStr(S, ".SetValueDisplay Row") Then Stop
'If IsInStr(S, "cmdSaleTotals.Move") Then Stop
'If IsInStr(S, "2830") Then Stop
'If IsInStr(S, "True") Then Stop
'If IsInStr(S, ":=") Then Stop
'If IsInStr(S, "GetRecordNotFound") Then Stop
'If IsInStr(S, "Nonretro_14day") Then Stop
'If IsInStr(S, "Git") Then Stop
'If IsInStr(S, "GitFolder") Then Stop
'If IsInStr(S, "Array") Then Stop

  S = RegExReplace(S, patNotToken & patToken & "!" & patToken & patNotToken, "$1$2(""$3"")$4") ' RS!Field -> RS("Field")
  S = RegExReplace(S, "^" & patToken & "!" & patToken & patNotToken, "$1(""$2"")$3") ' RS!Field -> RS("Field")

  S = RegExReplace(S, "([^a-zA-Z0-9_.])NullDate([^a-zA-Z0-9_.])", "$1NullDate()$2")
  
  S = ConvertVb6Specific(S, Complete)
  If Complete Then ConvertElement = S: Exit Function
  
  If RegExTest(Trim(S), "^" & patToken & "$") Then
'    If S = "SqFt" Then Stop
    If IsFuncRef(Trim(S)) And S <> CurrSub Then
      ConvertElement = Trim(S) & "()"
      Exit Function
    ElseIf IsPrivateFuncRef(CurrentModule, Trim(S)) And S <> CurrSub Then
      ConvertElement = Trim(S) & "()"
      Exit Function
    ElseIf IsEnumRef(Trim(S)) Then
      ConvertElement = EnumRefRepl(Trim(S))
      Exit Function
    End If
  End If

  If RegExTest(Trim(S), "^" & patTokenDot & "$") And StrCnt(S, ".") = 1 Then
'    If S = "SqFt" Then Stop
    Dim First As String, Second As String
    First = SplitWord(S, 1, ".")
    Second = SplitWord(S, 2, ".")
    If IsModuleRef(First) And IsFuncRef(Second) Then
      If IsFuncRef(Trim(Second)) And S <> CurrSub Then
        ConvertElement = Trim(S) & "()"
        Exit Function
      ElseIf IsEnumRef(Trim(S)) Then
        ConvertElement = EnumRefRepl(Trim(S))
        Exit Function
      End If
    End If
  End If

'If IsInStr(S, "Not optTagIncoming") Then Stop
  If IsControlRef(Trim(S), FormName) Then
'If IsInStr(S, "optTagIncoming") Then Stop
    S = FormControlRepl(S, FormName)
  ElseIf LMatch(Trim(S), "Not ") And IsControlRef(Mid(Trim(S), 5), FormName) Then
    S = "!(" & FormControlRepl(Mid(Trim(S), 5), FormName) & ")"
  End If
  
  If IsFormRef(Trim(S)) Then
    ConvertElement = FormRefRepl(Trim(S))
    Exit Function
  End If
  

  FirstToken = RegExNMatch(S, patTokenDot, 0)
  FirstWord = SplitWord(S, 1)
  If FirstWord = "Not" Then
    S = "!" & ConvertValue(Mid(S, 5))
    FirstWord = SplitWord(Mid(S, 2))
  End If
  If S = FirstWord Then ConvertElement = S: GoTo ManageFunctions
  If S = FirstToken Then ConvertElement = S & "()": GoTo ManageFunctions
  
  If FirstToken = FirstWord And Not isOperator(SplitWord(S, 2)) Then ' Sub without parenthesis
    ConvertElement = FirstWord & "(" & SplitWord(S, 2, , , True) & ")"
  Else
    ConvertElement = S
  End If
  
ManageFunctions:
'If IsInStr(ConvertElement, "New_CDbTypeAhead") Then Stop
  If RegExTest(ConvertElement, "(\!)?[a-zA-Z0-9_.]+[ ]*\(.*\)$") Then
    If (Left(ConvertElement, 1) = "!") Then
      ConvertElement = "!" & ConvertFunctionCall(Mid(ConvertElement, 2))
    Else
      ConvertElement = ConvertFunctionCall(ConvertElement)
    End If
  End If

DoReplacements:
  If IsInStr(ConvertElement, ":=") Then
    Dim Ts As String
    Ts = SplitWord(ConvertElement, 1, ":=")
    Ts = Ts & ": "
    Ts = Ts & ConvertElement(SplitWord(ConvertElement, 2, ":=", True, True))
    ConvertElement = Ts
  End If

  ConvertElement = Replace(ConvertElement, " & ", " + ")
  ConvertElement = Replace(ConvertElement, " = ", " == ")
  ConvertElement = Replace(ConvertElement, "<>", " != ")
  ConvertElement = Replace(ConvertElement, " Not ", " !")
  ConvertElement = Replace(ConvertElement, "(Not ", "(!")
  ConvertElement = Replace(ConvertElement, " Or ", " || ")
  ConvertElement = Replace(ConvertElement, " And ", " && ")
  ConvertElement = Replace(ConvertElement, " Mod ", " % ")
  ConvertElement = Replace(ConvertElement, "Err.", "Err().")
  ConvertElement = Replace(ConvertElement, "Debug.Print", "Console.WriteLine")
  
  ConvertElement = Replace(ConvertElement, "NullDate", "NullDate")
  Do While IsInStr(ConvertElement, ", ,")
    ConvertElement = Replace(ConvertElement, ", ,", ", _,")
  Loop
  ConvertElement = Replace(ConvertElement, "(,", "(_,")

'If IsInStr(ConvertElement, "&H") And Right(ConvertElement, 1) = "&" Then Stop
'If IsInStr(ConvertElement, "1/1/2001") Then Stop

  ConvertElement = RegExReplace(ConvertElement, "([0-9])#", "$1")
  
  If Left(ConvertElement, 2) = "&H" Then
    ConvertElement = "0x" & Mid(ConvertElement, 3)
    If Right(ConvertElement, 1) = "&" Then ConvertElement = Left(ConvertElement, Len(ConvertElement) - 1)
  End If
  
  If WithLevel > 0 Then
    T = Stack(WithVars, , True)
    ConvertElement = Trim(RegExReplace(ConvertElement, "([ (])(\.)" & patToken, "$1" & T & "$2$3"))
    If Left(ConvertElement, 1) = "." Then ConvertElement = T & ConvertElement
  End If
End Function

Public Function ConvertFunctionCall(ByVal fCall As String) As String
  Dim I As Long, N As Long, TB As String, Ts As String, TName As String
  Dim TV As String
  Dim vP As Variable
'Debug.Print "ConvertFunctionCall: " & fCall

  TB = ""
  TName = RegExNMatch(fCall, "^[a-zA-Z0-9_.]*")
  TB = TB & TName

  Ts = Mid(fCall, Len(TName) + 2)
  Ts = Left(Ts, Len(Ts) - 1)
  
  vP = SubParam(TName)
  If ConvertDataType(vP.asType) = "Recordset" Then
    TB = TB & ".Fields["
    TB = TB & ConvertValue(Ts)
    TB = TB & "].Value"
  ElseIf vP.asArray <> "" Then
    TB = TB & "["
    TB = TB & ConvertValue(Ts)
    TB = TB & "]"
'    TB = Replace(TB, ", ", "][")
  Else
    N = nextByPCt(Ts, ",")
    TB = TB & "("
    For I = 1 To N
      If I <> 1 Then TB = TB & ", "
      TV = nextByP(Ts, ",", I)
      If IsFuncRef(TName) Then
        If Trim(TV) = "" Then
          TB = TB & ConvertElement(FuncRefArgDefault(TName, I))
        Else
          If FuncRefArgByRef(TName, I) Then TB = TB & "ref "
          TB = TB & ConvertValue(TV)
        End If
      Else
        TB = TB & ConvertValue(TV)
      End If
    Next
    TB = TB & ")"
  End If
  ConvertFunctionCall = TB
End Function


Public Function ConvertValue(ByVal S As String) As String
  Dim F As String, Op As String, OpN As String
  Dim O As String
  O = ""
  S = Trim(S)
  If S = "" Then Exit Function
  
'If IsInStr(S, "GetMaxFieldValue") Then Stop
'If IsInStr(S, "DBAccessGeneral") Then Stop
'If IsInStr(S, "tallable") Then Stop
'If Left(S, 3) = "RS(" Then Stop
'If Left(S, 6) = "DBName" Then Stop
'If Left(S, 6) = "fName" Then Stop

  SubParamUsedList TokenList(S)
  
  If RegExTest(S, "^-[a-zA-Z0-9_]") Then
    ConvertValue = "-" & ConvertValue(Mid(S, 2))
    Exit Function
  End If
  
  Do While True
    F = NextByOp(S, 1, Op)
    If F = "" Then Exit Do
    Select Case Trim(Op)
      Case "\":    OpN = "/"
      Case "=":    OpN = " == "
      Case "<>":   OpN = " != "
      Case "&":    OpN = " + "
      Case "Mod":  OpN = " % "
      Case "Is":   OpN = " == "
      Case "Like": OpN = " == "
      Case "And":  OpN = " && "
      Case "Or":   OpN = " || "
      Case Else:   OpN = Op
    End Select
    
    
    If Left(F, 1) = "(" And Right(F, 1) = ")" Then
      O = O & "(" & ConvertValue(Mid(F, 2, Len(F) - 2)) & ")" & OpN
    Else
      O = O & ConvertElement(F) & OpN
    End If
    
    If Op = "" Then Exit Do
    S = Mid(S, Len(F) + Len(Op) + 1)
    If S = "" Or Op = "" Then Exit Do
  Loop
  ConvertValue = O
End Function

Public Function ConvertGlobals(ByVal Str As String, Optional ByVal AsModule As Boolean = False) As String
  Dim Res As String
  Dim S() As String, L As Variant, O As String
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
  InitDeString
  For Each L In S
    L = DeComment(L)
    L = DeString(L)
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
    ElseIf RegExTest(L, "^(Public |Private |)Declare ") Then
      O = ConvertAPIDef(L)
    ElseIf RegExTest(L, "^(Global |Public |Private |)Const ") Then
      O = ConvertConstant(L, True)
    ElseIf RegExTest(L, "^(Public |Private |)Event ") Then
      O = ConvertEvent(L)
    ElseIf RegExTest(L, "^(Public |Private |)Enum ") Then
      Building = L
    ElseIf RegExTest(LTrim(L), "^(Public |Private |)Type ") Then
      Building = L
    ElseIf tLeft(L, 8) = "Private " Or tLeft(L, 7) = "Public " Or tLeft(L, 4) = "Dim " Then
      O = ConvertDeclare(L, 0, True, AsModule)
    End If
      
    O = ReComment(O)
    Res = Res & ReComment(O) & IIf(O = "" Or Right(O, 2) = vbCrLf, "", vbCrLf)
    N = N + 1
'    Prg N
'    If N Mod 10000 = 0 Then Stop
  Next
'  Prg

  Res = ReString(Res, True)
  ConvertGlobals = Res
End Function

Public Function ConvertCodeLine(ByVal S As String) As String
  Dim T As Long, A As String, B As String
  Dim P As String, V As Variable
  Dim FirstWord As String, Rest As String
  Dim N As Long

'If IsInStr(S, "dbClose") Then Stop
'If IsInStr(S, "Nothing") Then Stop
'If IsInStr(S, "Close ") Then Stop
'If IsInStr(S, "& functionType & fieldInfo &") Then Stop
'If IsInStr(S, " & vbCrLf2 & Res)") Then Stop
'If IsInStr(S, "Res = CompareSI(SI1, SI2)") Then Stop
'If IsInStr(S, "frmPrintPreviewDocument") Then Stop
'If IsInStr(S, "NewAudit.Name1") Then Stop
'If IsInStr(S, "optDelivered") Then Stop
'If IsInStr(S, " Is Nothing Then") Then Stop
'If IsInStr(S, "SqFt, SqYd") Then Stop
'If IsInStr(S, "optTagIncoming") Then Stop
'If IsInStr(S, "Kill modAshleyItemAlign") Then Stop
'If IsInStr(S, "PRFolder") Then Stop
'If IsInStr(S, "Array()") Then Stop
'If IsInStr(S, "App.Path") Then Stop

  If Trim(S) = "" Then ConvertCodeLine = "": Exit Function
  Dim Complete As Boolean
  S = ConvertVb6Specific(S, Complete)
  If Complete Then
    ConvertCodeLine = S
    Exit Function
  End If
  
  If RegExTest(Trim(S), "^[a-zA-Z0-9_.()]+ \= ") Or RegExTest(Trim(S), "^Set [a-zA-Z0-9_.()]+ \= ") Then ' Assignment
    T = InStr(S, "=")
    A = Trim(Left(S, T - 1))
    If tLeft(A, 4) = "Set " Then A = Trim(Mid(A, 5))
    SubParamAssign RegExNMatch(A, patToken)
    If RegExTest(A, "^" & patToken & "\(""[^""]+""\)") Then
      P = RegExNMatch(A, "^" & patToken)
      V = SubParam(P)
      If V.Name = P Then
        SubParamAssign P
        Select Case V.asType
          Case "Recordset", "ADODB.Recordset"
            ConvertCodeLine = RegExReplace(A, "^" & patToken & "(\("")([^""]+)(""\))", "$1.Fields[""$3""].Value")
          Case Else
            If Left(A, 1) = "." Then A = Stack(WithVars, , True) & A
            ConvertCodeLine = A
        End Select
      End If
    Else
      If Left(A, 1) = "." Then A = Stack(WithVars, , True) & A
      ConvertCodeLine = A
    End If
    
    Dim tAWord As String
    tAWord = SplitWord(A, 1, ".")
    If IsFormRef(tAWord) Then
      A = Replace(A, tAWord, tAWord & ".instance", , 1)
    End If
    
    ConvertCodeLine = ConvertValue(ConvertCodeLine) & " = "

    B = ConvertValue(Trim(Mid(S, T + 1)))
    ConvertCodeLine = ConvertCodeLine & B
  Else
'Debug.Print S
'If IsInStr(S, "Call ") Then Stop
    If LMatch(LTrim(S), "Call ") Then S = Mid(LTrim(S), 6)

    FirstWord = SplitWord(Trim(S))
    Rest = SplitWord(Trim(S), 2, , , True)
    If Rest = "" Then
      ConvertCodeLine = S & IIf(Right(S, 1) <> ")", "()", "")
      ConvertCodeLine = ConvertElement(ConvertCodeLine)
    ElseIf FirstWord = "RaiseEvent" Then
      ConvertCodeLine = ConvertValue(S)
    ElseIf FirstWord = "Debug.Print" Then
      ConvertCodeLine = "Console.WriteLine(" & ConvertValue(Rest) & ")"
    ElseIf StrQCnt(FirstWord, "(") = 0 Then
      ConvertCodeLine = ""
      ConvertCodeLine = ConvertCodeLine & FirstWord & "("
      N = 0
      Do
        N = N + 1
        B = nextByP(Rest, ", ", N)
        If B = "" Then Exit Do
        ConvertCodeLine = ConvertCodeLine & IIf(N = 1, "", ", ") & ConvertValue(B)
      Loop While True
      ConvertCodeLine = ConvertCodeLine & ")"
'      ConvertCodeLine = ConvertElement(ConvertCodeLine)
    Else
      ConvertCodeLine = ConvertValue(S)
    End If
    If WithLevel > 0 And Left(Trim(ConvertCodeLine), 1) = "." Then ConvertCodeLine = Stack(WithVars, , True) & Trim(ConvertCodeLine)
  End If
  
'  If IsInStr(ConvertCodeLine, ",,,,,,,") Then Stop
  
  ConvertCodeLine = ConvertCodeLine & ";"
'Debug.Print ConvertCodeLine
End Function

Public Function PostConvertCodeLine(ByVal Str As String) As String
  Dim S As String
  S = Str
  
'  If IsInStr(S, "optPoNo") Then Stop
  
  If IsInStr(S, "0 &") Then S = Replace(S, "0 &", "0")
  If IsInStr(S, ".instance.instance") Then S = Replace(S, ".instance.instance", ".instance")
  If IsInStr(S, ".IsChecked)") Then S = Replace(S, ".IsChecked)", ".IsChecked == true)", 1)
  If IsInStr(S, ".IsChecked &") Then S = Replace(S, ".IsChecked", ".IsChecked == true", 1)
  If IsInStr(S, ".IsChecked |") Then S = Replace(S, ".IsChecked", ".IsChecked == true", 1)
  If IsInStr(S, ".IsChecked,") Then S = Replace(S, ".IsChecked", ".IsChecked == true", 1)
  If IsInStr(S, ".IsChecked == 1,") Then S = Replace(S, ".IsChecked == 1", ".IsChecked == true", 1)
  If IsInStr(S, ".IsChecked == 0,") Then S = Replace(S, ".IsChecked == 1", ".IsChecked == false", 1)
  
  If IsInStr(S, ".Visibility = true") Then S = Replace(S, ".Visibility = true", ".setVisible(true)")
  If IsInStr(S, ".Visibility = false") Then S = Replace(S, ".Visibility = false", ".setVisible(false)")
  
  If IsInStr(S, ".Print(") Then
    If IsInStr(S, ";);") Then
      S = Replace(S, ";);", ");")
      S = Replace(S, "Print(", "PrintNNL(")
    End If
    S = Replace(S, "; ", ", ")
  End If
  If IsInStr(S, ".Line((") Then
    S = Replace(S, ") - (", ", ")
    S = Replace(S, "Line((", "Line(")
    S = Replace(S, "));", ");")
  End If
  
  S = Replace(S, "vbRetryCancel +", "vbRetryCancel |")
  S = Replace(S, "vbOkOnly +", "vbOkOnly |")
  S = Replace(S, "vbOkCancel +", "vbOkCancel |")
  S = Replace(S, "vbExclamation +", "vbExclamation |")
  S = Replace(S, "vbYesNo +", "vbYesNo |")
  S = Replace(S, "vbQuestion +", "vbQuestion |")
  S = Replace(S, "vbOKCancel +", "vbOKCancel |")
  S = Replace(S, "+ vbExclamation", "| vbExclamation")
  
  PostConvertCodeLine = S
End Function

Public Function ConvertSub(ByVal Str As String, Optional ByVal AsModule As Boolean = False, Optional ByVal ScanFirst As VbTriState = vbUseDefault) As String
  Dim oStr As String
  Dim Res As String
  Dim S() As String, L As Variant, O As String, T As String, U As String, V As String
  Dim CM As Long, cN As Long
  Dim K As Long
  Dim Ind As Long
  Dim inCase As Long
  Dim returnVariable As String
  
'  If IsInStr(Str, "Dim oFTP As New FTP") Then Stop
'  If IsInStr(Str, "cHolding") Then Stop
'If IsInStr(Str, "IsIDE") Then Stop

  
  Select Case ScanFirst
    Case vbUseDefault:
      oStr = Str
      ConvertSub oStr, AsModule, vbTrue
'                          If IsInStr(Str, "StoreStockToolTipText") Then Stop
      ConvertSub = ConvertSub(oStr, AsModule, vbFalse)
      Exit Function
    Case vbTrue:        SubBegin
    Case vbFalse:       SubBegin True
  End Select
  

  
  Res = ""
  Str = Replace(Str, vbLf, "")
  S = Split(Str, vbCr)
  Ind = 0
    
'If IsInStr(Str, " WinCDSDataPath(") Then Stop
'If IsInStr(Str, " RunShellExecute(") Then Stop
'If IsInStr(Str, " ValidateSI(") Then Stop
  For Each L In S
'If IsInStr(L, "OrdVoid") Then Stop
'If IsInStr(L, "MsgBox") Then Stop
'If IsInStr(L, "And Not IsDoddsLtd Then") Then Stop
    L = DeComment(L)
    L = DeString(L)
    O = ""

'If IsInStr(L, "1/1/2001") Then Stop
'If ScanFirst = vbFalse Then Stop
'If IsInStr(L, "Public Function GetFileAutonumber") Then Stop
'If IsInStr(L, "GetCustomerBalance") Then Stop
'If IsInStr(L, "IsIDE") Then Stop


    Dim PP As String, PQ As String
    PP = "^(Public |Private |)(Friend |)(Function |Sub )" & patToken & "[ ]*\("
    PQ = "^(Public |Private )(Property )(Get |Let |Set )" & patToken & "[ ]*\("
    If RegExNMatch(L, PP) <> "" Then
      Dim nK As Long
'      CurrSub = nextBy(L, "(", 1)
'      If (LMatch(CurrSub, "Public ")) Then CurrSub = Mid(CurrSub, 8)
'      If (LMatch(CurrSub, "Private ")) Then CurrSub = Mid(CurrSub, 9)
'      If (LMatch(CurrSub, "Friend ")) Then CurrSub = Mid(CurrSub, 8)
'      If (LMatch(CurrSub, "Function ")) Then CurrSub = Mid(CurrSub, 10)
'      If (LMatch(CurrSub, "Sub ")) Then CurrSub = Mid(CurrSub, 5)
'If IsInStr(L, "Public Function IsIn") Then Stop
      O = O & sSpace(Ind) & ConvertPrototype(L, returnVariable, AsModule, CurrSub)
      Ind = Ind + SpIndent
    ElseIf RegExNMatch(L, PQ) <> "" Then
'      If IsInStr(L, "edi888_Admin888_Src") Then Stop
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
    ElseIf tLMatch(L, "GoTo ") Then
      O = O & "goto " & SplitWord(Trim(L), 2) & ";"
    ElseIf RegExTest(Trim(L), "^[a-zA-Z_][a-zA-Z_0-9]*:$") Then ' Goto Label
      O = O & L & ";" ' c# requires a trailing ; on goto labels without trailing statements.  Likely a C# bug/oversight, but it's there.
    ElseIf tLeft(L, 3) = "Dim" Then
      O = ConvertDeclare(L, Ind)
    ElseIf tLeft(L, 5) = "Const" Then
      O = sSpace(Ind) & ConvertConstant(L, False)
    ElseIf tLeft(L, 3) = "If " Then  ' Code sanitization prevents all single-line ifs.
'If IsInStr(L, "optDelivered") Then Stop
'If IsInStr(L, "PRFolder") Then Stop
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
      O = O & "break;" & vbCrLf
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
        CM = Val(SplitWord(T, 2, " To "))
        For K = cN To CM
          O = O & "case " & K & ": "
        Next
      Else
        Dim TT As Variant, LL As Variant
'          O = O & sSpace(Ind) & "case " & ConvertValue(T) & ":"
        O = O & Space(Ind)
        For Each LL In Split(T, ",")
          O = O & "case " & ConvertValue(T) & ": "
        Next
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
      Dim iterVar As String
      iterVar = SplitWord(L, 1, " In ")
      O = O & sSpace(Ind) & "foreach(var iter" & iterVar & " in " & SplitWord(L, 2, " In ") & ") {" & vbCrLf & iterVar & " = iter" & iterVar & ";"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 4) = "For " Then
      Dim forKey As String, forStr As String, forEnd As String
      L = tMid(L, 5)
      forKey = SplitWord(L, 1, "=")
      L = SplitWord(L, 2, "=")
      forStr = SplitWord(L, 1, " To ")
      forEnd = SplitWord(L, 2, " To ")
      O = O & sSpace(Ind) & "for(" & ConvertElement(forKey) & "=" & ConvertElement(forStr) & "; " & ConvertElement(forKey) & "<" & ConvertElement(forEnd) & "; " & ConvertElement(forKey) & "++) {"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 11) = "Loop While " Then
      Ind = Ind - SpIndent
      O = O & sSpace(Ind) & "} while(!(" & ConvertValue(tMid(L, 12)) & "));"
    ElseIf tLeft(L, 11) = "Loop Until " Then
      Ind = Ind - SpIndent
      O = O & sSpace(Ind) & "} while(!(" & ConvertValue(tMid(L, 12)) & "));"
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

      T = ConvertValue(tMid(L, 6))
      U = ConvertDataType(SubParam(T).asType)
      V = WithMark & IIf(SubParam(T).Name <> "", T, Random)
      If U = "" Then U = DefaultDataType
      
      Stack WithAssign, T
      Stack WithTypes, U
      Stack WithVars, V
      
      O = O & sSpace(Ind) & U & " " & V & ";" & vbCrLf
      MaxWithLevel = MaxWithLevel + 1
      O = O & sSpace(Ind) & V & " = " & T & ";"
      Ind = Ind + SpIndent
    ElseIf tLeft(L, 8) = "End With" Then
      WithLevel = WithLevel - 1
      T = Stack(WithAssign)
      U = Stack(WithTypes)
      V = Stack(WithVars)
      If SubParam(T).Name <> "" Then
        O = O & sSpace(Ind) & T & " = " & V & ";"
      End If
      Ind = Ind - SpIndent
    ElseIf IsInStr(L, "On Error ") Or IsInStr(L, "Resume ") Then
      O = sSpace(Ind) & "// TODO (not supported): " & L
    Else
'If IsInStr(L, "ComputeAgeing dtpArrearControlDate") Then Stop
'If IsInStr(L, "RaiseEvent") Then Stop
'If IsInStr(L, "Debug.Print") Then Stop
'If IsInStr(L, "HasGit") Then Stop
      O = sSpace(Ind) & ConvertCodeLine(L)
    End If
    
    O = modOrigConvert.PostConvertCodeLine(O)
    O = modProjectSpecific.ProjectSpecificPostCodeLineConvert(O)
    
    O = ReComment(O)
    Res = Res & ReComment(O) & IIf(O = "", "", vbCrLf)
  Next
  
  ConvertSub = Res
End Function

