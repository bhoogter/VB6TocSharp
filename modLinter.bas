Attribute VB_Name = "modLinter"
Option Explicit
':::: modLinter
':::SUMMARY
': Lint VB6 files
':
':::DESCRIPTION
': Inspect VB6 files to linting.  I.e., format, spacing, and other code-quality features.
':
'::: Rules
': - Indentation
': - File Names
':
':::SEE ALSO
':    - modXML

Public LintForBuild As Boolean

Private Const lintFileShort_Len As Long = 20
Private Const lintLint_MaxErrors As Long = 30
Private Const lintLint_TabWidth As Long = 2
Private Const lintLint_IndentContextDiv As String = ":"
Private Const lintLint_MaxBlankLines As Long = 5
Private Const lintLint_MaxBlankLines_AtClose As Long = 2
Private Const lintNoLint_ScanRange As Long = 10
Private Const lintTag_ScanRange As Long = 10
Private Const lintTag_Key As String = "@"
Private Const lintTag_Start As String = "'" & lintTag_Key
Private Const lintTag_Div As String = "-"
Private Const lintTag_NoLint As String = lintTag_Start & "NO-LINT"
Private Const lintFile_Option As String = "Option "
Private Const lintFile_Option_Explicit As String = "Explicit"
Private Const lintDotsPerRow As Long = 60
Private Const lintFixList_Sep As String = "/||\"
Private Const lintFixList_Div As String = ":oo:"

Public Enum lintErrorTypes
  ltUnkn
  ltLErr
  ltIdnt
  ltDECL
  ltDEPR
  ltSTOP
  ltWITH
  ltVarN
  ltArgN
  ltFunN
  ltCtlN
  ltObjN
  ltSelf
  ltType
  ltWhtS
  ltOptn
  ltBadC
  ltNTyp
  ltNOpD
End Enum

Private Function CheckNoLint(ByVal FileName As String, Optional ByVal lType As lintErrorTypes = ltUnkn, Optional ByVal vLine As String) As Boolean
  Dim I As Long, L As String, A As Long
  Dim CA As String, cP As String, cB As String
  CheckNoLint = False
  If LintAbbr(lType) = "" Then CheckNoLint = True: Exit Function  ' Disable lint type
  
  CA = lintTag_NoLint
  cP = lintTag_NoLint & lintTag_Div
  cB = lintTag_NoLint & lintTag_Div & UCase(LintAbbr(lType))
  
  If IsInStr(vLine, Mid(CA, 2)) And Not IsInStr(vLine, Mid(cP, 2)) Then CheckNoLint = True: Exit Function
  If IsInStr(vLine, Mid(cB, 2)) Then CheckNoLint = True: Exit Function
  
  A = LintModuleFirstLine(FileName)
  For I = A To A + lintTag_ScanRange
    L = UCase(ReadFile(FileName, I, 1))
    If lType = ltUnkn Then
      If LMatch(L, CA) And Not LMatch(L, cP) Then
        CheckNoLint = True
        Exit Function
      End If
    Else
      If LMatch(L, cB) Then
        CheckNoLint = True
        Exit Function
      End If
    End If
  Next
End Function

Private Function LintAbbr(ByVal lType As lintErrorTypes, Optional ByRef TypeName As String) As String
  ' if this function returns "", the lint type is ignored.  Add the following after the normal lint type to disable:
  ' : LintAbbr = ""
  Select Case lType
    Case ltLErr: LintAbbr = "LErr": TypeName = "Lint Runtime Error"
    Case ltIdnt: LintAbbr = "Idnt": TypeName = "Indent"
    Case ltDECL: LintAbbr = "Decl": TypeName = "Declaration"
    Case ltDEPR: LintAbbr = "Depr": TypeName = "Deprecated"
    Case ltSTOP: LintAbbr = "STOP": TypeName = "Stop Encountered"
    Case ltWITH: LintAbbr = "WITH": TypeName = "With Statement" ': LintAbbr = IIf(Not LintForBuild, LintAbbr, "")
    Case ltVarN: LintAbbr = "VarN": TypeName = "Variable Name"
    Case ltArgN: LintAbbr = "ArgN": TypeName = "Argument Name"
    Case ltFunN: LintAbbr = "FunN": TypeName = "Function Name"
    Case ltCtlN: LintAbbr = "CtlN": TypeName = "Control Name"
    Case ltObjN: LintAbbr = "ObjN": TypeName = "Object Name"
    Case ltSelf: LintAbbr = "Self": TypeName = "Self Reference"
    Case ltType: LintAbbr = "Type": TypeName = "Data Type"
    Case ltWhtS: LintAbbr = "WhtS": TypeName = "White Space"
    Case ltOptn: LintAbbr = "Optn": TypeName = "Option"
    Case ltBadC: LintAbbr = "BadC": TypeName = "Bad Code"
    
    Case ltNTyp: LintAbbr = "NTyp": TypeName = "No Type" ': LintAbbr = IIf(Not LintForBuild, LintAbbr, "")
    Case ltNOpD: LintAbbr = "NOpD": TypeName = "No Default For Optional": LintAbbr = ""
    
    Case Else:   LintAbbr = "UNKN": TypeName = "Unknown"
  End Select
End Function
Private Function LintName(ByVal lType As lintErrorTypes) As String:   LintAbbr lType, LintName: End Function
Private Function LintFileShort(ByVal FFile As String) As String
  LintFileShort = AlignString(FileName(FFile), lintFileShort_Len)
End Function

Private Function AddErrStr(ByRef ErrStr As String, ByVal FileName As String, ByVal LineNo As String, ByVal vLine As String, ByVal Msg As String, ByVal lType As lintErrorTypes) As String
  Static ErrCnt As Long
  If CheckNoLint(FileName, lType, vLine) Then Exit Function
  
  If ErrStr = "" Then ErrCnt = 0
  ErrCnt = ErrCnt + 1
  If ErrCnt > lintLint_MaxErrors Then
    If Right(ErrStr, 4) <> " >>>" Then
      ErrStr = ErrStr & vbCrLf & "<<< Max Error Count Exceeded >>>"
    End If
    Exit Function
  End If
  If ErrStr <> "" Then ErrStr = ErrStr & vbCrLf
  ErrStr = ErrStr & LintFileShort(FileName) & " (Line " & LineNo & "): " & LintAbbr(lType) & " - " & Msg
End Function

Private Function AddIndent(ByRef Lvl As Long, ByRef Str As String, Optional ByRef Context As String = "", Optional ByVal POP As Boolean = False) As Boolean
  AddIndent = True
  Context = Replace(Context, lintLint_IndentContextDiv, "-")
  If POP Then
    Lvl = Lvl - lintLint_TabWidth
    If Lvl < 0 Then
      Lvl = 0
      Str = ""
      AddIndent = False
    End If
    Context = SplitWord(Str, -1, lintLint_IndentContextDiv)
    Str = Left(Str, Len(Str) - Len(Context))
    If Right(Str, Len(lintLint_IndentContextDiv)) = lintLint_IndentContextDiv Then Str = Left(Str, Len(Str) - Len(lintLint_IndentContextDiv))
  Else
    Lvl = Lvl + lintLint_TabWidth
    Str = Str & IIf(Str = "", "", lintLint_IndentContextDiv) & Context
  End If
End Function
Private Function IndentContext(ByVal Str As String) As String:    IndentContext = SplitWord(Str, -1, lintLint_IndentContextDiv): End Function

Private Function DeComment(ByVal S As String) As String
  Dim I As Long
  Dim C As String
  Dim Q As Boolean
  Q = False
  DeComment = S
  If IsNotInStr(S, "'") Then Exit Function
  
  For I = 1 To Len(S)
    C = Mid(S, I, 1)
    If C = """" Then Q = Not Q
    If Not Q And C = "'" Then
      DeComment = RTrim(Left(S, I - 1))
      Exit Function
    End If
  Next
End Function
Private Function DeString(ByVal S As String) As String
  Const Q As String = """"
  Const Token As String = "_"
  Dim A As Long, B As Long
  DeString = S
  A = InStr(S, Q)
  If A > 0 Then
    B = InStr(A + 1, S, Q)
    If B > 0 Then
      DeString = DeString(Left(S, A - 1) & Token & Mid(S, B + 1))
      Exit Function
    End If
  End If
  DeString = S
End Function
Private Function DeSpace(ByVal S As String) As String
  Dim N As Long
  DeSpace = S
  Do
    N = Len(DeSpace)
    DeSpace = Replace(DeSpace, "  ", " ")
    If Len(DeSpace) = N Then Exit Function
  Loop While True
End Function



Public Function LintFolder(Optional ByVal Folder As String, Optional ByVal AutoFix As Boolean = False, Optional ByVal ForBuild As Boolean) As Boolean
  LintForBuild = True
  LintFolder = LintFileList(VBPModules(vbpFile) & vbCrLf & VBPClasses(vbpFile) & vbCrLf & VBPForms(), AutoFix)
End Function

Public Function LintModules(Optional ByVal Folder As String, Optional ByVal AutoFix As Boolean = False) As Boolean
  LintModules = LintFileList(VBPModules, AutoFix)
End Function

Public Function LintClasses(Optional ByVal Folder As String, Optional ByVal AutoFix As Boolean = False) As Boolean
  LintClasses = LintFileList(VBPClasses, AutoFix)
End Function

Public Function LintForms(Optional ByVal Folder As String, Optional ByVal AutoFix As Boolean = False) As Boolean
  LintForms = LintFileList(VBPForms(), AutoFix)
End Function

Public Function LintFileList(ByVal List As String, ByVal AutoFix As Boolean) As Boolean
  Dim E As String, L As Variant
  Dim X As Long
  Dim StartTime As Date
  StartTime = Now
  
  For Each L In Split(List, vbCrLf)
    If Not LintFile(L, E) Then
      If AutoFix Then
        LintFileIndent DevelopmentFolder & L, , True
        Debug.Print "x";
      Else
        Debug.Print vbCrLf & "LINT FAILED: " & LintFileShort(L)
        MsgBox E, , "Lint Folder"
        Debug.Print E
        Debug.Print "?LintFile(""" & L & """)"
        Exit Function
      End If
    Else
      Debug.Print Switch(Right(L, 3) = "frm", "o", Right(L, 3) = "cls", "x", True, ".");
    End If
    X = X + 1
    If X >= lintDotsPerRow Then X = 0: Debug.Print
    DoEvents
  Next
  Debug.Print vbCrLf & "Done (" & DateDiff("s", StartTime, Now) & "secs)."
  LintFileList = True
End Function

Public Function LintFile(ByVal FileName As String, Optional ByRef ErrStr As String = "#", Optional ByVal AutoFix As Boolean = False) As Boolean
  Dim Alert As Boolean, aOutput As Boolean
  Alert = ErrStr = "#"
  aOutput = ErrStr = "."
  ErrStr = ""
  LintFile = True
  
'  FileName = MakePathAbsolute(FileName, DevelopmentFolder)
  If Not FileExists(FileName) Then LintFile = True: Exit Function
  If CheckNoLint(FileName) Then LintFile = True: Exit Function
  
  LintFile = LintFile And LintFileOptions(FileName, ErrStr)
  LintFile = LintFile And LintFileIndent(FileName, ErrStr, AutoFix)
  LintFile = LintFile And LintFileNaming(FileName, ErrStr, AutoFix)
  LintFile = LintFile And LintFileControlNaming(FileName, ErrStr, AutoFix)
  LintFile = LintFile And LintFileBadCode(FileName, ErrStr, AutoFix)

  If AutoFix Then                             ' Re-run to test after Auto-fix
    ErrStr = ""
    LintFile = LintFile(FileName, ErrStr)
  End If
  
  If ErrStr <> "" Then
    If aOutput Then Debug.Print ErrStr
    If Alert Then MsgBox ErrStr
  Else
    LintFile = True
  End If
End Function

Private Function LintModuleFirstLine(ByVal FileName As String) As Long
  Dim S As String, N As Long, K As String
  S = ReadEntireFile(FileName)
  S = Left(S, InStr(S, "Attribute VB_Name"))
  LintModuleFirstLine = CountLines(S, False, False)
  
  Do
    K = ReadFile(FileName, LintModuleFirstLine, 1)
    If Not LMatch(K, "Attribute ") Then Exit Function
    If K = "" Then Exit Function
    LintModuleFirstLine = LintModuleFirstLine + 1
  Loop While True
End Function

Public Function LintFileOptions(ByVal FileName As String, Optional ByRef ErrStr As String) As Boolean
  Dim I As Long, L As String, A As Long, F As String
  Dim oExplicit As Boolean
  
  LintFileOptions = True
  
  A = LintModuleFirstLine(FileName)
  For I = A To A + lintTag_ScanRange
    L = ReadFile(FileName, I, 1)
    If LMatch(L, lintFile_Option) Then
      F = Mid(L, Len(lintFile_Option) + 1)
      If F = lintFile_Option_Explicit Then
        oExplicit = True
      Else
        AddErrStr ErrStr, FileName, I - A + 1, L, "Prohibited Flag: Option " & F, ltOptn
        LintFileOptions = False
      End If
    End If
  Next
  
  If Not oExplicit Then
    AddErrStr ErrStr, FileName, 1, "", "Missing Flag: Option Explicit", ltOptn
    LintFileOptions = False
  End If
End Function

Private Function AutoFixInit(ByVal FileName As String) As String
  Dim A As Long, FL As String
  A = LintModuleFirstLine(FileName)
  AutoFixInit = DevelopmentFolder & "templint.txt"
  FL = ReadFile(FileName, 1, A - 1)
  WriteFile AutoFixInit, FL, True
End Function
Private Function AutoFixLine(ByVal FixFile As String, ByVal Line As String, ByVal LineFixes As String) As String
  Dim FixL As Variant, KSpl As Variant
  
  AutoFixLine = Line
  If LineFixes <> "" Then
    For Each FixL In Split(LineFixes, lintFixList_Sep)
      KSpl = Split(FixL, lintFixList_Div)
      If KSpl(0) = "^" Then AutoFixLine = KSpl(1) & AutoFixLine
      If KSpl(0) = "$" Then AutoFixLine = AutoFixLine & KSpl(1)
  '    If KSpl(0) = "#" Then Exit Function ' suppress output
      AutoFixLine = Replace(AutoFixLine, KSpl(0), KSpl(1))
    Next
  End If
  WriteFile FixFile, AutoFixLine
End Function
Private Function AddLineFixes(ByVal LineFixes As String, ByVal Find As String, ByVal Repl As String) As String
  AddLineFixes = LineFixes & IIf(Len(LineFixes) = 0, "", lintFixList_Sep) & Find & lintFixList_Div & Repl
End Function

Private Sub AutoFixFinalize(ByVal FileName As String, ByVal FixFile As String)
  Dim Contents As String
  Contents = ReadEntireFileAndDelete(FixFile)
  Do While Right(Contents, 1) = vbLf Or Right(Contents, 1) = vbCr
    Contents = Left(Contents, Len(Contents) - 1)
  Loop
  Contents = Contents & vbCrLf
  WriteFile FileName, Contents, True
End Sub

Public Function LintFileIndent(ByVal FileName As String, Optional ByRef ErrStr As String, Optional ByVal AutoFix As Boolean = False) As Boolean
  Dim A As Long
  Dim N As Long, I As Long
  Dim Continued As Long
  Dim Idnt As Long, Context As String
  Dim OL As String, L As String, LNo As Long, tL As String, FL As String
  Dim Blanks As Long
  Dim FixFile As String
  Dim LineFixes As String
  
  If Not FileExists(FileName) Then LintFileIndent = True: Exit Function
On Error GoTo FailedLint
  
  N = CountFileLines(FileName)
  A = LintModuleFirstLine(FileName)
  If AutoFix Then FixFile = AutoFixInit(FileName)
  
  For I = A To N
    L = ReadFile(FileName, I, 1)
    OL = L
    FL = L
    If Trim(L) = "" Then Blanks = Blanks + 1
    L = DeComment(L)
    tL = LTrim(L)
    LineFixes = ""
    If LMatch(L, "Attribute ") Then GoTo NotRealLine
    LNo = I - A + 1
'    If IsDevelopment And LNo > 275 Then Stop
    
    If Trim(L) = "" Then
      If Blanks = lintLint_MaxBlankLines + 1 Then AddErrStr ErrStr, FileName, LNo, OL, "Too many sequential blank lines.", ltWhtS
      GoTo SkipLine
    End If
    If Continued Then GoTo SkipLine
    
    Blanks = 0
    
    If Len(L) = Len(tL) And Right(L, 1) = ":" Then GoTo SkipLine    ' Goto Marks
    If LMatch(tL, "On Error ") Then GoTo SkipLine                   ' Error Handlers
    If LMatch(tL, "Debug.") Then GoTo SkipLine                      ' Error Handlers
    If LMatch(tL, "ActiveLog") Then GoTo SkipLine                   ' Active Logging
    If Left(tL, 1) = "#" Then GoTo SkipLine                         ' Processer Directives
    
    If LMatch(tL, "End Select") Then
      If IndentContext(Context) = "Select Case Item" Then AddIndent Idnt, Context, , True
      AddIndent Idnt, Context, , True
    ElseIf LMatch(tL, "End ") Or _
       LMatch(tL, "ElseIf ") Or _
       LMatch(tL, "Else") And Not LMatch(tL, "Else ") Or _
        IsIn(IndentContext(Context), "For Loop", "For Each Loop") And tL = "Next" Or LMatch(tL, "Next ") Or _
        IndentContext(Context) = "Do While Loop" And LMatch(tL, "Loop") Or _
        IndentContext(Context) = "Do Until Loop" And LMatch(tL, "Loop") Or _
        IndentContext(Context) = "Do Loop" And LMatch(tL, "Loop") _
        Then
      If Not AddIndent(Idnt, Context, , True) Then
        AddErrStr ErrStr, FileName, LNo, OL, "Cannot set negative indent.", ltIdnt
      End If
    ElseIf LMatch(tL, "Case ") Then
      If IndentContext(Context) = "Select Case Item" Then AddIndent Idnt, Context, , True
    End If
    
'If LNo >= 383 Then Stop
'If InStr(FileName, "Functions") Then Stop
'If IsInStr(tL, "Property") Then Stop
    If Idnt <> (Len(L) - Len(tL)) Then
      AddErrStr ErrStr, FileName, LNo, OL, "Expected Indent " & Idnt & ", is " & (Len(L) - Len(tL)) & ": " & IndentContext(Context), ltIdnt
      FL = Space(Idnt) & LTrim(OL)
    End If
    
    If LMatch(DeString(tL), "Declare ") Then
      ' ignore API declarations for now
    ElseIf LMatch(tL, "Function ") Then
      AddErrStr ErrStr, FileName, LNo, OL, "Function should be declared either Public or Private.  Neither specified.", ltDECL
      If IsNotInStr(DeSpace(L), ": End ") Then AddIndent Idnt, Context, "Function"
      LineFixes = AddLineFixes(LineFixes, "^", "Public ")
    ElseIf LMatch(tL, "Sub ") Then
      AddErrStr ErrStr, FileName, LNo, OL, "Sub should be declared either Public or Private.  Neither specified.", ltDECL
      If IsNotInStr(DeSpace(L), ": End ") Then AddIndent Idnt, Context, "Sub"
      LineFixes = AddLineFixes(LineFixes, "^", "Public ")
    ElseIf LMatch(tL, "Property ") Then
      AddErrStr ErrStr, FileName, LNo, OL, "Property should be declared either Public or Private.  Neither specified.", ltDECL
      If IsNotInStr(DeSpace(L), ": End ") Then AddIndent Idnt, Context, "Property"
      LineFixes = AddLineFixes(LineFixes, "^", "Public ")
    ElseIf LMatch(tL, "Private Function ") Or LMatch(tL, "Private Sub ") Or LMatch(tL, "Private Property ") _
        Or LMatch(tL, "Public Function ") Or LMatch(tL, "Public Sub ") Or LMatch(tL, "Public Property ") _
        Or LMatch(tL, "Friend Function ") Or LMatch(tL, "Friend Sub ") Or LMatch(tL, "Friend Property ") _
        Then
      If IsNotInStr(DeSpace(L), ": End ") Then AddIndent Idnt, Context, SplitWord(tL, 2)
    ElseIf LMatch(tL, "For Each ") Then
      If IsNotInStr(DeSpace(L), ": Next") Then AddIndent Idnt, Context, "For Each Loop"
    ElseIf LMatch(tL, "For ") Then
      If IsNotInStr(DeSpace(L), ": Next") Then AddIndent Idnt, Context, "For Loop"
    ElseIf LMatch(tL, "While ") Then
      If IsNotInStr(DeSpace(L), ": Loop") Then AddIndent Idnt, Context, "While Loop"
    ElseIf tL = "Do" Then
      If IsNotInStr(DeSpace(L), ": Loop") Then AddIndent Idnt, Context, "Do Loop"
    ElseIf LMatch(tL, "Do While ") Then
      If IsNotInStr(DeSpace(L), ": Loop") Then AddIndent Idnt, Context, "Do While Loop"
    ElseIf LMatch(tL, "Do Until ") Then
      If IsNotInStr(DeSpace(L), ": Loop") Then AddIndent Idnt, Context, "Do Until Loop"
    ElseIf LMatch(tL, "With ") Then
      AddErrStr ErrStr, FileName, LNo, OL, "WITH Deprecated--unsupported in all upgrade paths.", ltWITH
      If IsNotInStr(L, "End With") Then AddIndent Idnt, Context, "With Block"
    ElseIf LMatch(tL, "Select Case ") Then
      AddIndent Idnt, Context, "Select Block"
    ElseIf LMatch(tL, "Case ") Then
'      If IndentContext(Context) = "Select Case Item" Then AddIndent Idnt, Context, , True
      If IsNotInStr(tL, ": ") Then AddIndent Idnt, Context, "Select Case Item"
    ElseIf (LMatch(tL, "Type ") Or LMatch(tL, "Private Type ") Or LMatch(tL, "Public Type ")) And Not LMatch(tL, "Type As ") Then
      If IsNotInStr(L, "End Type") Then AddIndent Idnt, Context, "Type Def"
    ElseIf LMatch(tL, "Enum ") Or LMatch(tL, "Private Enum") Or LMatch(tL, "Public Enum") Then
      If IsNotInStr(L, "End Enum") Then AddIndent Idnt, Context, "Enum"
    ElseIf LMatch(tL, "If ") Then
      If Right(tL, 5) = " Then" Or Right(tL, 2) = " _" Then AddIndent Idnt, Context, "If Block"
    ElseIf LMatch(tL, "Else") And Not LMatch(tL, "Else ") Then
      AddIndent Idnt, Context, "Else Block"
    ElseIf LMatch(tL, "ElseIf ") Then
      AddIndent Idnt, Context, "ElseIf Block"
    End If
    
    If IsInStr(DeString(tL), "Wend") Then
      AddErrStr ErrStr, FileName, LNo, OL, "WEND is deprecated.  Use Do While X ... Loop or Do ... Loop While X", ltDEPR
    ElseIf IsInStr(" " & DeString(tL), " Next ") Then
      AddErrStr ErrStr, FileName, LNo, OL, "NEXT no longer needs its operand.  Remove Variable name after Next.", ltDEPR
    ElseIf IsInStr(" " & DeString(tL), " Call ") Then
      AddErrStr ErrStr, FileName, LNo, OL, "CALL is no longer required.  Do not use CALL keyword in code.", ltDEPR
      LineFixes = AddLineFixes(LineFixes, "Call ", "")
    ElseIf IsInStr(DeString(tL), "GoSub") Then
      AddErrStr ErrStr, FileName, LNo, OL, "GOSUB is deprecated and should not be used.", ltDEPR
    ElseIf IsInStr(DeString(tL), "$(") Then
      AddErrStr ErrStr, FileName, LNo, OL, "Type-casting functions is deprecated.  Remove $ before (...).", ltDEPR
      LineFixes = AddLineFixes(LineFixes, "$(", "(")
    ElseIf tL = "Return" Then
      AddErrStr ErrStr, FileName, LNo, OL, "GOSUB / RETURN is deprecated and should not be used.", ltDEPR
    ElseIf IsInStr(DeString(tL), " Stop") And Right(tL, 4) = "Stop" Then
      If Not IsInStr(tL, "IsDevelopment") Then
        AddErrStr ErrStr, FileName, LNo, OL, "Code contains STOP statement.", ltSTOP
      End If
    End If
  
SkipLine:
    Continued = (Right(L, 2) = " _")

NotRealLine:
    If AutoFix Then AutoFixLine FixFile, FL, LineFixes
  Next
  
  If Idnt <> 0 Then
    AddErrStr ErrStr, FileName, LNo, OL, "Indent did not close. EOF.", ltIdnt
  End If
  If Blanks > lintLint_MaxBlankLines_AtClose Then
    AddErrStr ErrStr, FileName, LNo, OL, "Too many blank lines at end of file.  Max=" & lintLint_MaxBlankLines_AtClose & ".", ltWhtS
  End If
  
  If AutoFix Then AutoFixFinalize FileName, FixFile
  
  Exit Function
  
FailedLint:
  AddErrStr ErrStr, FileName, LNo, "", "Lint Error", ltLErr
  Resume Next
End Function

Private Function LintFileTestName(ByVal dName As String, ByRef ErrStr As String) As Boolean
  LintFileTestName = False
  
  ' TODO: This check is only to avoid the problem of Dim SomVar(0, 0) for commas embedded in var names...
  If ReduceString(dName, STR_CHR_UCASE & STR_CHR_LCASE & "_", "", , False) = "" Then LintFileTestName = True: Exit Function
  
  If dName = LCase(dName) Then
    ErrStr = "Name [" & dName & "] Is All Lower Case"
  ElseIf IsIn(Right(dName, 1), "%", "&", "@", "!", "#", "$") Then
    Dim C As String, TName As String
'% Integer Dim L%
'& Long  Dim M&
'@ Decimal Const W@ = 37.5
'! Single  Dim Q!
'# Double  Dim X#
'$ String  Dim V$ = "Secret"
    C = Right(dName, 1)
    TName = Switch(C = "%", "Long", C = "&", "Long", C = "@", "Double", C = "!", "Double", C = "#", "Double", C = "$", "String", True, "UNKNOWN-TYPE-KEY")
    ErrStr = "Type declaration by variable name not allowed. Replace " & Right(dName, 1) & " with type " & TName & "."
  Else
    ErrStr = ""
    LintFileTestName = True
  End If
End Function

Private Function LintStandardNaming(ByVal vN As String) As String
  Select Case LCase(vN)
' Capitalize All
    Case "nl"
      LintStandardNaming = UCase(vN)
' Capitalize Second Letter...
    Case "vn", "dx", "dy", "dname", "vdata", "tstr"
      LintStandardNaming = LCase(Left(vN, 1)) & UCase(Mid(vN, 2, 1)) & LCase(Mid(vN, 3))
' Capitalize First Letter (default)
    Case Else
      LintStandardNaming = Capitalize(vN)
  End Select
End Function

Private Function LintFileTestArgN(ByVal dName As String, ByRef ErrStr As String) As Boolean
  LintFileTestArgN = LintFileTestName(dName, ErrStr)
End Function

Private Function LintFileTestType(ByVal DType As String, ByRef ErrStr As String) As Boolean
  LintFileTestType = True
  Select Case Trim(DType)
    Case "Integer"
      ErrStr = "Integer should not be used here.  Use Long."
      LintFileTestType = False
'    Case "Single"
'      ErrStr = "Single should not be used here.  Use Double."
'      LintFileTestType = False
    Case "Short"
      LintFileTestType = False
      ErrStr = "Short should not be used here.  Use Long."
  End Select
End Function

Private Function LintFileIsEvent(ByVal fName As String, ByVal tL As String) As Boolean
  LintFileIsEvent = False
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_") And IsInStr(tL, "Private ")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_Click")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_DblClick")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_KeyDown")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_KeyUp")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_KeyPress")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_KeyDown")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_Validate")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_GotFocus")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_LostFocus")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_QueryUnload")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_OLEDragDrop")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_OLESetData")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_RouteAfterCalculate")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_Error")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_DataArrival")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_Connect")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_Close")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_ConnectionRequest")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_SendComplete")
  LintFileIsEvent = LintFileIsEvent Or IsInStr(fName, "_ZipThreadDone")
End Function

Private Function LintFileNaming(ByVal FileName As String, Optional ByRef ErrStr As String, Optional ByVal AutoFix As Boolean = False) As Boolean
  Dim LNo As String
  Dim A As Long, N As Long, I As Long, tE As String
  Dim OL As String, L As String, tL As String
  Dim fName As String, vArgs As String, AName As String, vDef As String
  Dim isLet As Boolean, isSet As Boolean
  Dim vRetType As String
  Dim vName As String, vType As String
  Dim Continued As Boolean
  Dim Decl As Variant
  Dim FixFile As String, LineFixes As String
  
  If AutoFix Then FixFile = AutoFixInit(FileName)
  
  N = CountFileLines(FileName)
  A = LintModuleFirstLine(FileName)
  For I = A To N
    OL = ReadFile(FileName, I, 1)
    L = DeComment(OL)
    tL = LTrim(L)
    LNo = I - A + 1
    LineFixes = ""
    If Continued Then GoTo SkipLine
'    If IsDevelopment And LNo > 1822 Then Stop
'    If LNo = 58 Then Stop
'If LNo = 147 Then Stop
    
    
    If LMatch(tL, "Public Function ") Or _
            LMatch(tL, "Public Sub ") Or _
            LMatch(tL, "Public Property ") Or _
            LMatch(tL, "Private Function ") Or _
            LMatch(tL, "Private Sub ") Or _
            LMatch(tL, "Private Property ") Or _
            LMatch(tL, "Friend Function ") Or _
            LMatch(tL, "Friend Sub ") Or _
            LMatch(tL, "Friend Property ") Or _
            LMatch(tL, "Function ") Or _
            LMatch(tL, "Sub ") Or _
            LMatch(tL, "Property ") _
            Then
      '
      fName = SplitWord(tL, 1, "(")
      fName = Replace(fName, "Public ", "")
      fName = Replace(fName, "Private ", "")
      fName = Replace(fName, "Friend ", "")
      fName = Replace(fName, "Function ", "")
      fName = Replace(fName, "Sub ", "")
      fName = Replace(fName, "Property ", "")
      fName = Replace(fName, "Get ", "")
      isLet = Left(Trim(fName), 4) = "Let "
      fName = Replace(fName, "Let ", "")
      isSet = Left(Trim(fName), 4) = "Set "
      fName = Replace(fName, "Set ", "")
      
      If Not LintFileTestName(fName, tE) Then
        AddErrStr ErrStr, FileName, LNo, OL, tE, ltVarN
        LineFixes = AddLineFixes(LineFixes, " " & tE, " " & LintStandardNaming(tE))
      End If
      
'If fName = "Form_QueryUnload" Then Stop
      If Not LintFileIsEvent(fName, tL) Then
        vRetType = SplitWord(tL, 2, ")")
        If Left(vRetType, 3) = "As " Then
          vRetType = Mid(vRetType, 4)
          If Not LintFileTestType(vRetType, tE) Then
            AddErrStr ErrStr, FileName, LNo, OL, tE, ltType
          End If
        Else
          If IsNotInStr(OL, "Sub ") And Right(OL, 1) <> "_" And Not isLet And Not isSet Then
            AddErrStr ErrStr, FileName, LNo, OL, "No Return Type On Func/Prop", ltNTyp
          End If
        End If
        vArgs = SplitWord(DeString(tL), 1, ":")
        vArgs = SplitWord(vArgs, 2, "(", , True)
        Dim MM As Long
        If vArgs <> "" Then
          MM = IIf(Right(vArgs, 2) = "()", InStrRev(vArgs, ")", Len(vArgs) - 2), InStrRev(vArgs, ")")) - 1
        End If
        If MM >= 0 Then vArgs = Left(vArgs, MM)
        For Each Decl In Split(DeString(vArgs), ",")
          Decl = Trim(Decl)
          If Decl = "_" Then GoTo IgnoreParam               ' Not checking multi-line declarations for now..  Could insert in-place multi-line read..
          
          If LMatch(Decl, "Optional ") Then
            vDef = SplitWord(Decl, 2, " = ")
            If vDef = "" Then
              AddErrStr ErrStr, FileName, LNo, OL, "Parameter declared OPTIONAL but no default specified. Must specify default.", ltNOpD
            End If
            Decl = Trim(Replace(Decl, "Optional ", ""))
          End If
          
          If Not LMatch(Decl, "ByVal ") And Not LMatch(Decl, "ByRef ") And Not LMatch(Decl, "ParamArray ") Then
            AddErrStr ErrStr, FileName, LNo, OL, "Neither ByVal nor ByRef are specified. Must Specify one or other.", ltDECL
            LineFixes = AddLineFixes(LineFixes, Replace(Decl, "_", ""), "ByRef " & Replace(Decl, "_", ""))
          Else
            Decl = Replace(Decl, "ByRef ", "")
            Decl = Replace(Decl, "ByVal ", "")
            Decl = Replace(Decl, "ParamArray ", "")
            Decl = Trim(Decl)
          End If
          
          vName = SplitWord(Decl, 1, " As ")
          If Not LintFileTestArgN(vName, tE) Then
            AddErrStr ErrStr, FileName, LNo, OL, tE, ltArgN
          End If
          
          vType = SplitWord(Decl, 2, " As ")
          If vType = "" Then
            AddErrStr ErrStr, FileName, LNo, OL, "No Param Type on Func/Sub/Prop", ltNTyp
          End If
          If Not LintFileTestType(vType, tE) Then
            AddErrStr ErrStr, FileName, LNo, OL, tE, ltType
          End If
          
IgnoreParam:
        Next
      End If
    ElseIf LMatch(tL, "Private Declare ") Or _
            LMatch(tL, "Public Declare ") Or _
            LMatch(tL, "Declare ") Then
      '
    ElseIf LMatch(tL, "Dim ") Or _
            LMatch(tL, "Private ") Or _
            LMatch(tL, "Public ") Then
      
      vArgs = tL
      vArgs = Replace(vArgs, "Dim ", "")
      vArgs = Replace(vArgs, "Private ", "")
      vArgs = Replace(vArgs, "Public ", "")
      vArgs = Replace(vArgs, "Const ", "")
      
      For Each Decl In Split(DeString(vArgs), ",")
        vName = Trim(SplitWord(Decl, 1, " As "))
        vName = Trim(SplitWord(vName, 1, " = "))
        If Not LintFileTestName(vName, tE) Then
          AddErrStr ErrStr, FileName, LNo, OL, tE, ltArgN
          LineFixes = AddLineFixes(LineFixes, vName, LintStandardNaming(vName))
        End If
        If IsNotInStr(OL, "Enum ") And IsNotInStr(OL, "Type ") Then
          vType = Trim(SplitWord(Decl, 2, " As "))
          If Not LMatch(vName, "Event ") Then
            If vType = "" Then AddErrStr ErrStr, FileName, LNo, OL, "No Type on Decl", ltNTyp
          End If
          If Not LintFileTestType(vType, tE) Then
            AddErrStr ErrStr, FileName, LNo, OL, tE, ltType
          End If
        End If
      Next
    End If
SkipLine:
    Continued = (Right(L, 2) = " _")

    If AutoFix Then AutoFixLine FixFile, OL, LineFixes
  Next
  
  If AutoFix Then AutoFixFinalize FileName, FixFile
  
  LintFileNaming = ErrStr = ""
End Function

Private Function LintFileControlNaming(ByVal FileName As String, Optional ByRef ErrStr As String, Optional ByVal AutoFix As Boolean = False) As Boolean
  Const MaxCtrl As Long = 128
  Dim LNo As Long
  Dim Contents As String, I As Long
  Dim Match As String
  Dim N As Long, K As Long
  Dim CtlName As String, ErrMsg As String
  Dim cUnique As Collection, Reported As Variant
  
  Contents = ReadEntireFile(FileName)
  Set cUnique = New Collection
  
  Dim vTypes() As Variant
  vTypes = Array("CheckBox", "Command", "Option", "Frame", "Label", "TextBox", "RichTextBox", "RichTextBoxNew", "ComboBox", "ListBox", "Timer", "UpDown", "HScrollBar", "Image", "Picture", "MSFlexGrid", "DBGrid", "Line", "Shape", "DTPicker")
  For I = LBound(vTypes) To UBound(vTypes)
   ' "[^a-zA-Z]" & vTypes(I) & "[0-9*]\."
    Match = "Begin [a-zA-Z0-9]*.[a-zA-Z0-9]* " & vTypes(I) & "[0-9]*"
    If RegExTest(Contents, Match) Then
      N = RegExCount(Contents, Match)
      For K = 0 To N - 1
        CtlName = RegExNMatch(Contents, Match, K)
        CtlName = Split(CtlName, " ")(2)
        CtlName = Trim(CtlName)

        On Error Resume Next
        Reported = ""
        Reported = cUnique.Item(CtlName)
        cUnique.Add "1", CtlName
        On Error GoTo 0
      
        If CtlName <> "" And Reported = "" Then
          ErrMsg = "Default Control Name in use: " & CtlName & ".  Rename Control."
          AddErrStr ErrStr, FileName, LNo, "", ErrMsg, ltCtlN
        End If
      Next
    End If
  Next
  
  LintFileControlNaming = ErrStr = ""
End Function

Public Function LintFileBadCode(ByVal FileName As String, Optional ByRef ErrStr As String, Optional ByVal AutoFix As Boolean = False) As Boolean
  Dim LNo As String
  Dim A As Long, N As Long, I As Long, tE As String
  Dim OL As String, L As String, tL As String
  Dim fName As String, vArgs As String, AName As String, vDef As String
  Dim vRetType As String
  Dim vName As String, vType As String
  Dim Continued As Boolean
  Dim Decl As Variant
  Dim FixFile As String, LineFixes As String
  
  If AutoFix Then FixFile = AutoFixInit(FileName)
  
  N = CountFileLines(FileName)
  A = LintModuleFirstLine(FileName)
  For I = A To N
    OL = ReadFile(FileName, I, 1)
    L = DeComment(OL)
    tL = LTrim(L)
    LNo = I - A + 1
    LineFixes = ""
    If Continued Then GoTo SkipLine
    
    If RegExTest(tL, "\.Enabled = [-0-9]") Then AddErrStr ErrStr, FileName, LNo, OL, "Property [Enabled] Should Be Boolean.  Numeric found.", ltType
    If RegExTest(tL, "\.Visible = [-0-9]") Then AddErrStr ErrStr, FileName, LNo, OL, "Property [Visible] Should Be Boolean.  Numeric found.", ltType
    If RegExTest(" " & tL, "[^a-zA-Z0-0]Me[.][^ ]") Then AddErrStr ErrStr, FileName, LNo, OL, "Self Reference [Me.*] is unnecessary.", ltSelf '@NO-LINT
  
SkipLine:
    Continued = (Right(L, 2) = " _")
  Next

  LintFileBadCode = ErrStr = ""
End Function
