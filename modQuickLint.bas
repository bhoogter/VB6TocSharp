Attribute VB_Name = "modQuickLint"
Option Explicit
 
Private Const Idnt As Long = 2
Private Const MAX_ERRORS_DEFAULT As Long = 50
Private Const Attr As String = "Attribute"
Private Const Q As String = """"
Private Const A As String = "'"
Private Const S As String = " "
Private Const LintKey As String = "'@NO-LINT"

Private Const TY_ALLTY As String = "AllTy"

Private Const TY_ERROR As String = "Error"
Private Const TY_INDNT As String = "Indnt"
Private Const TY_ARGNA As String = "ArgNa"
Private Const TY_ARGTY As String = "ArgTy"
Private Const TY_FSPNA As String = "FSPNa"
Private Const TY_DEPRE As String = "Depre"
Private Const TY_MIGRA As String = "Migra"
Private Const TY_STYLE As String = "Style"
Private Const TY_BLANK As String = "Blank"
Private Const TY_EXPLI As String = "Expli"
Private Const TY_COMPA As String = "Compa"
Private Const TY_TYPEC As String = "TypeC"
Private Const TY_NOTYP As String = "NoTyp"
Private Const TY_BYRFV As String = "ByReV"
Private Const TY_PRIPU As String = "PriPu"
Private Const TY_FNCRE As String = "FncRe"
Private Const TY_CORRE As String = "Corre"
Private Const TY_GOSUB As String = "GoSub"
Private Const TY_CSTOP As String = "CStop"
Private Const TY_OPDEF As String = "OpDef"
Private Const TY_OPBYR As String = "OpByR"
Private Const TY_DFCTL As String = "DfCtl"

Private Const DISABLED_LINT_TYPES As String = TY_OPBYR ' TY_ARGTY & "," & TY_OPDEF
Private Const WARNING_LINT_TYPES As String = ""
Private Const AUTOFIX_LINT_TYPES As String = TY_INDNT & "," & TY_ARGNA & "," & TY_OPDEF & "," & TY_NOTYP & "," & TY_STYLE

Public ErrorPrefix As String
Public ErrorIgnore As String
Public AutofixFind() As String
Public AutofixRepl() As String
Public AutofixFindRestOfFile() As String
Public AutofixReplRestOfFile() As String

Public WellKnownNames As New Collection

Public Function ErrorTypes() As Variant()
  ErrorTypes = Array(TY_ALLTY, TY_ERROR, TY_INDNT, TY_ARGNA, TY_ARGTY, TY_FSPNA, TY_DEPRE, TY_MIGRA, TY_STYLE, TY_BLANK, TY_EXPLI, TY_COMPA, TY_TYPEC, TY_NOTYP, TY_BYRFV, TY_PRIPU, TY_FNCRE, TY_CORRE, TY_GOSUB, TY_CSTOP, TY_OPDEF, TY_OPBYR, TY_DFCTL)
End Function

Private Function ResolveSources(ByVal FileName As String) As String
  If FileName = "" Then FileName = "prj.vbp"
  If FileName = "forms" Then
    ResolveSources = VBPForms(True)
  ElseIf FileName = "modules" Then
    ResolveSources = VBPModules
  ElseIf FileName = "classes" Then
    ResolveSources = VBPClasses
  ElseIf FileName = "usercontrols" Then
    ResolveSources = VBPUserControls
  Else
    If InStr(FileName, "\") = 0 Then FileName = App.Path & "\" & FileName
    ResolveSources = IIf(Right(FileName, 4) = ".vbp", VBPCode(FileName), FileName)
  End If
End Function

Public Function Lint(Optional ByVal FileName As String = "", Optional ByVal Alert As Boolean = False) As String
  Dim FileList As String
  FileList = ResolveSources(FileName)
  Lint = QuickLintFiles(FileList, MAX_ERRORS_DEFAULT)
  If Alert And Lint <> "" Then MsgBox Lint
End Function

Public Function LintFix(Optional ByVal FileName As String = "") As String
  Dim FileList As String
  FileList = ResolveSources(FileName)
  QuickLintFiles FileList, 0, True
End Function

Public Function QuickLintFiles(ByVal List As String, Optional ByVal MaxErrors As Long = MAX_ERRORS_DEFAULT, Optional ByVal AutoFix As Boolean = False) As String
  Const lintDotsPerRow As Long = 50
  Dim L As Variant
  Dim X As Long
  Dim StartTime As Date
  StartTime = Now
  
  For Each L In Split(List, vbCrLf)
    Dim Result As String
    Result = QuickLintFile(L, MaxErrors, AutoFix)
    If Not Result = "" Then
      Dim S As String
      Debug.Print vbCrLf & "Done (" & DateDiff("s", StartTime, Now) & "s).   To re-run for failing file, hit enter on the line below:"
      S = "LINT FAILED: " & L & vbCrLf & Result & vbCrLf & "?Lint(""" & L & """)"
      QuickLintFiles = S
      Exit Function
    Else
      Debug.Print Switch(Right(L, 3) = "frm", "o", Right(L, 3) = "cls", "x", True, ".");
    End If
    X = X + 1
    If X >= lintDotsPerRow Then X = 0: Debug.Print
    DoEvents
  Next
  Debug.Print vbCrLf & "Done (" & DateDiff("s", StartTime, Now) & "s)."
  QuickLintFiles = ""
End Function

Public Function QuickLintFile(ByVal File As String, Optional ByVal MaxErrors As Long = MAX_ERRORS_DEFAULT, Optional ByVal AutoFix As Boolean = False) As String
  If InStr(File, "\") = 0 Then File = App.Path & "\" & File
  Dim fName As String, Contents As String, GivenName As String, CheckName As String
  fName = Mid(File, InStrRev(File, "\") + 1)
  CheckName = Replace(Replace(Replace(fName, ".bas", ""), ".cls", ""), ".frm", "")
  ErrorPrefix = Right(Space(18) & fName, 18) & " "
  Contents = ReadEntireFile(File)
  GivenName = RegExNMatch(Contents, "Attribute VB_Name = ""([^""]+)""", 0)
  GivenName = Replace(Replace(GivenName, "Attribute VB_Name = ", ""), """", "")
  If LCase(CheckName) <> LCase(GivenName) Then
    QuickLintFile = "Module name [" & GivenName & "] must match file name [" & fName & "].  Rename module or class to match the other"
    Exit Function
  End If
  QuickLintFile = QuickLintContents(Contents, MaxErrors, IIf(AutoFix, File, ""))
End Function

Public Function QuickLintContents(ByVal Contents As String, Optional ByVal MaxErrors As Long = MAX_ERRORS_DEFAULT, Optional ByVal AutoFix As String = "") As String
  Dim Lines() As String, ActualLine As Variant, LL As String, L As String
On Error GoTo LintError
  ErrorIgnore = DISABLED_LINT_TYPES
  Lines = Split(Replace(Contents, vbCr, ""), vbLf)
  Erase AutofixFind
  Erase AutofixRepl
  Erase AutofixFindRestOfFile
  Erase AutofixReplRestOfFile

  Dim InAttributes As Boolean, InBody As Boolean
    
  Dim MultiLineOrig As String, MultiLine As String, IsMultiLine As Boolean
  Dim Indent As Long, LineN As Long
  Dim Errors As String, ErrorCount As Long
  Dim BlankLineCount As Long
  Dim Options As New Collection
  Dim NewContents As String

  Indent = 0
  
  TestDefaultControlNames Errors, ErrorCount, 0, Contents
  
  For Each ActualLine In Lines
    LL = ActualLine
    If MaxErrors > 0 And ErrorCount >= MaxErrors Then Exit For
    
    IsMultiLine = False
    If Right(LL, 2) = " _" Then
      Dim Portion As String
      Portion = Left(LL, Len(LL) - 2)
      MultiLineOrig = MultiLineOrig & LL & vbCrLf
      If MultiLine <> "" Then Portion = " " & Trim(Portion)
      MultiLine = MultiLine + Portion
      LineN = LineN + 1
      GoTo NextLineWithoutRecord
    ElseIf MultiLine <> "" Then
      MultiLineOrig = MultiLineOrig & LL
      LL = MultiLine & " " & Trim(LL)
      MultiLine = ""
      IsMultiLine = True
    Else
      MultiLineOrig = ""
    End If
    
    TestBlankLines Errors, ErrorCount, LineN, LL, BlankLineCount
    TestLintControl LL
    L = CleanLine(LL)
    
    If Not InBody Then
      Dim IsAttribute As Boolean
      IsAttribute = Left(L, 10) = "Attribute "
      If Not InAttributes And IsAttribute Then
        InAttributes = True
        GoTo NextLine
      ElseIf InAttributes And Not IsAttribute Then
        InAttributes = False
        InBody = True
        LineN = 0
      Else
        GoTo NextLine
      End If
    End If
    
    LineN = LineN + 1
'    If LineN = 487 Then Stop
    
    Dim UnindentedAlready As Boolean
    If RegExTest(L, " ^Option ") Then Options.Add "true", Replace(L, "Options ", "")
    
    If RegExTest(L, "^[ ]*(Else|ElseIf .* Then)$") Then
      Indent = Indent - Idnt
    ElseIf RegExTest(L, "^[ ]*End Select$") Then
      Indent = Indent - Idnt - Idnt
    ElseIf RegExTest(L, "^[ ]*(End (If|Function|Sub|Property|Enum|Type)|Next( .*)?|Wend|Loop|Loop (While .*|Until .*)|ElseIf .*)$") Then
      Indent = Indent - Idnt
      UnindentedAlready = True
    Else
      UnindentedAlready = False
    End If
    
    Dim LineIndent As Long
    LineIndent = 0
    Do While Mid(RTrim(L), LineIndent + 1, 1) = S: LineIndent = LineIndent + 1: Loop
    TestIndent Errors, ErrorCount, LineN, L, LineIndent, IIf(Not RegExTest(L, "^[ ]*Case "), Indent, Indent - Idnt)
    
    Dim Statements() As String, SS As Variant, St As String
    Statements = Split(L, ": ")
    For Each SS In Statements
      St = SS
      
      If RegExTest(L, "^[ ]*(Else|ElseIf .*)$") Then
        Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*(End (If|Function|Sub|Property)|Loop|Loop .*|Enum|Type|Select)$") Then
        If Not UnindentedAlready Then Indent = Indent - Idnt
      ElseIf RegExTest(St, "^[ ]*If ") Then
        If Not RegExTest(St, "\bThen ") Then Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*For ") Then
        Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*Next$") Then
        If Not UnindentedAlready Then Indent = Indent - Idnt
      ElseIf RegExTest(St, "^[ ]*Next [a-zA-Z_][a-zA-Z0-9_]*$") Then
        RecordError Errors, ErrorCount, TY_STYLE, LineN, "Remove variable from NEXT statement"
        AddFix TY_STYLE, "Next [a-zA-Z_][a-zA-Z0-9_]*$", "Next"
        If Not UnindentedAlready Then Indent = Indent - Idnt
      ElseIf RegExTest(St, "^[ ]*While ") Then
        RecordError Errors, ErrorCount, TY_STYLE, LineN, "Use Do While/Until...Loop in place of While...Wend"
        AddFix TY_STYLE, "\bWhile\b", "Do While"
        Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*Wend") Then
        AddFix TY_STYLE, "\bWend\b", "Loop"
      ElseIf RegExTest(St, "^[ ]*Do (While|Until)") Then
        Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*Loop$") Then
      ElseIf RegExTest(St, "^[ ]*Do$") Then
        Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*Loop While") Then
        Indent = Indent - Idnt
      ElseIf RegExTest(St, "^[ ]*Select Case ") Then
        Indent = Indent + Idnt + Idnt
      ElseIf RegExTest(St, "^[ ]*With ") Then
        RecordError Errors, ErrorCount, TY_MIGRA, LineN, "Remove all uses of WITH.  No migration path exists."
      ElseIf RegExTest(St, "^[ ]*(Private |Public )?Declare (Function |Sub )") Then
        ' External Api
      ElseIf RegExTest(St, "^((Private|Public|Friend) )?Function ") Then
        If Not RegExTest(St, ": End Function") Then Indent = Indent + Idnt
        TestSignature Errors, ErrorCount, LineN, St
      ElseIf RegExTest(St, "^((Private|Public|Friend) )?Sub ") Then
        If Not RegExTest(St, ": End Sub") Then Indent = Indent + Idnt
        TestSignature Errors, ErrorCount, LineN, St
      ElseIf RegExTest(St, "^((Private|Public|Friend) )?Property (Get|Let|Set) ") Then
        If Not RegExTest(St, ": End Property") Then Indent = Indent + Idnt
        TestSignature Errors, ErrorCount, LineN, St
      ElseIf RegExTest(St, "^[ ]*(Public |Private )?(Enum |Type )") Then
        Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*(Public |Private )?Declare ") Then
        Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*(Dim|Private|Public|Const|Global) ") Then
        TestDeclaration Errors, ErrorCount, LineN, St, False
      Else
        TestCodeLine Errors, ErrorCount, LineN, St
      End If
NextStatement:
    Next
NextLine:
    If AutoFix <> "" Then
      Dim Fixed As String
'      If IsMultiLine Then Stop
'      If InStr(LL, "Function") > 0 Then Stop
'      If InStr(LL, "Private Function") > 0 Then Stop
      If IsMultiLine Then
        Fixed = PerformAutofix(MultiLineOrig)
      Else
        Fixed = PerformAutofix(LL)
      End If
      NewContents = NewContents & Fixed & vbCrLf
    End If
NextLineWithoutRecord:
  Next
  If AutoFix <> "" Then WriteFile AutoFix, Left(NewContents, Len(NewContents) - 2), True
  
  TestModuleOptions Errors, ErrorCount, Options
  
  QuickLintContents = Errors
  Exit Function
LintError:
  RecordError Errors, ErrorCount, TY_ERROR, 0, "Linter Error [" & Err.Number & "]: " & Err.Description
  QuickLintContents = Errors
End Function

Private Function ReadEntireFile(ByVal tFileName As String) As String
On Error Resume Next
  Dim mFSO As Object
  Set mFSO = CreateObject("Scripting.FileSystemObject")
  ReadEntireFile = mFSO.OpenTextFile(tFileName, 1).ReadAll
  
  If FileLen(tFileName) / 10 <> Len(ReadEntireFile) / 10 Then
    MsgBox "ReadEntireFile was short: " & FileLen(tFileName) & " vs " & Len(ReadEntireFile)
  End If
End Function

Public Function CleanLine(ByVal Line As String) As String
  Dim X As Long, Y As Long
  Do While True
    X = InStr(Line, Q)
    If X = 0 Then Exit Do
    
    Y = InStr(X + 1, Line, Q)
    Do While Mid(Line, Y + 1, 1) = Q
      Y = InStr(Y + 2, Line, Q)
    Loop
    
    If Y = 0 Then Exit Do
    Line = Left(Line, X - 1) & String(Y - X + 1, "S") & Mid(Line, Y + 1)
  Loop
  
  X = InStr(Line, A)
  If X > 0 Then Line = RTrim(Left(Line, X - 1))
  
  CleanLine = Line
End Function
  
Public Sub RecordError(ByRef Errors As String, ByRef ErrorCount As Long, ByVal Typ As String, ByVal LineN As Long, ByVal Error As String)
  Dim eLine As String
  If InStr(UCase(ErrorIgnore), UCase(Typ)) > 0 Or InStr(ErrorIgnore, TY_ALLTY) > 0 Then Exit Sub

  If Len(Errors) <> 0 Then Errors = Errors & vbCrLf
  If InStr(Join(ErrorTypes, ","), Typ) = 0 Then
    Errors = Errors & ErrorPrefix & "[" & TY_ERROR & "] Line " & Right(Space(5) & LineN, 5) & ": Unknown error type in linter (add to ErrorTypes): " & Typ
  End If
  eLine = ErrorPrefix & "[" & Right(Space(5) & Typ, 5) & "] Line " & Right(Space(5) & LineN, 5) & ": " & Error
  If InStr(WARNING_LINT_TYPES, Typ) > 0 Then
    Debug.Print "WARNING: " & eLine
  Else
    Errors = Errors & eLine
    ErrorCount = ErrorCount + 1
  End If
End Sub

Public Function StartsWith(ByVal L As String, ByVal Find As String) As Boolean
  StartsWith = Left(L, Len(Find)) = Find
End Function

Public Function StripLeft(ByVal L As String, ByVal Find As String) As String
  If StartsWith(L, Find) Then StripLeft = Mid(L, Len(Find) + 1) Else StripLeft = L
End Function

Public Function RecordLeft(ByRef L As String, ByVal Find As String) As Boolean
  RecordLeft = StartsWith(L, Find)
  If RecordLeft Then L = Mid(L, Len(Find) + 1)
End Function

Public Sub TestIndent(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal L As String, ByVal LineIndent As Long, ByVal ExpectedIndent As Long)
  If RTrim(L) = "" Then Exit Sub
  If RegExTest(L, "^On Error ") Then Exit Sub
  If RegExTest(L, "^[a-zA-Z][a-zA-Z0-9]*:$") Then Exit Sub
  If RegExTest(L, "#(If|End If|Else|Const)") Then Exit Sub
  If StartsWith(L, "Debug.") Then Exit Sub
    
  If LineIndent <> ExpectedIndent Then
    RecordError Errors, ErrorCount, TY_INDNT, LineN, "Incorrect Indent -- expected " & ExpectedIndent & ", got " & LineIndent
    AddFix TY_INDNT, "^[ ]*", Space(ExpectedIndent)
  End If
End Sub

Public Sub TestBlankLines(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal L As String, ByRef BlankLineCount As Long)
  If Trim(L) <> "" Then
    BlankLineCount = 0
    Exit Sub
  End If
  BlankLineCount = BlankLineCount + 1
  If BlankLineCount > 3 Then RecordError Errors, ErrorCount, TY_BLANK, LineN, "Too many blank lines."
End Sub

Public Sub TestLintControl(ByVal L As String)
  Dim LL As Variant
  If InStr(L, LintKey) = 0 Then Exit Sub
  Dim Match As String, Typ As String
  Match = RegExNMatch(L, LintKey & "(-.....)?", 0)
  Typ = IIf(Match = LintKey, TY_ALLTY, Replace(Match, LintKey & "-", ""))
  ErrorIgnore = ErrorIgnore & "," & Typ
End Sub

Public Sub TestModuleOptions(ByRef Errors As String, ByRef ErrorCount As Long, ByVal Options As Collection)
On Error Resume Next
  Dim Value As String
  Value = ""
  Value = Options("Explicit")
  If Value <> "" Then RecordError Errors, ErrorCount, TY_EXPLI, 0, "Option Explicit not set on file"

  Value = ""
  Value = Options("Compare Binary")
  Value = Options("Compare Database")
  If Value <> "" Then RecordError Errors, ErrorCount, TY_COMPA, 0, "Use of Option Compare not recommended"
End Sub

Public Sub TestArgName(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal Name As String)
  Dim LL As String
  LL = Trim(Name)
  
  If RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*%$") Then ' % Integer Dim L%
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Integer deprecated: " & LL
    LL = Left(LL, Len(LL) - 1)
    AddFix TY_TYPEC, "\b" & LL & ".\b", LL & " As Integer"
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*&$") Then ' & Long  Dim M&
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Long deprecated: " & LL
    LL = Left(LL, Len(LL) - 1)
    AddFix TY_TYPEC, "\b" & LL & ".\b", LL & " As Long"
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*@$") Then ' @ Decimal Const W@ = 37.5
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Decimal deprecated: " & LL
    LL = Left(LL, Len(LL) - 1)
    AddFix TY_TYPEC, "\b" & LL & ".\b", LL & " As Decimal"
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-TY_TYPEC-Z0-9_]*!$") Then ' ! Single  Dim Q!
    RecordError Errors, ErrorCount, TY_DEPRE, LineN, "Use of Type Character For Single deprecated: " & LL
    LL = Left(LL, Len(LL) - 1)
    AddFix TY_TYPEC, "\b" & LL & ".\b", LL & " As Single"
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*#$") Then ' # Double  Dim X#
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Double deprecated: " & LL
    LL = Left(LL, Len(LL) - 1)
    AddFix TY_TYPEC, "\b" & LL & ".\b", LL & " As Double"
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*\$$") Then ' $ String  Dim V$ = "Secret"
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For String deprecated: " & LL
    LL = Left(LL, Len(LL) - 1)
    AddFix TY_TYPEC, "\b" & LL & ".\b", LL & " As String"
  End If
  
  If RegExTest(LL, "^[a-z][a-z0-9_]*[%&@!#$]?$") Then
    RecordError Errors, ErrorCount, TY_ARGNA, LineN, "Identifier name declared as all lower-case: " & LL
    AddFix TY_ARGNA, "\b" & LL & "\b", WellKnownName(LL), True
  End If
End Sub

Public Function WellKnownName(ByVal Str As String) As String
On Error Resume Next
  InitWellKnownNames
  WellKnownName = ""
  WellKnownName = WellKnownNames(LCase(Str))
  If WellKnownName = "" Then WellKnownName = Capitalize(Str)
End Function

Private Sub AddWellKnownName(ByVal S As String)
On Error Resume Next
  WellKnownNames.Add S, LCase(S)
End Sub

Public Sub InitWellKnownNames()
  Dim L As Variant
  If WellKnownNames.Count > 0 Then Exit Sub
  For Each L In Array("hWnd")
    AddWellKnownName L
  Next
End Sub

Public Sub TestSignatureName(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal Name As String)
  Dim LL As String
  LL = Trim(Name)
  
  If RegExTest(LL, "^[a-z][a-z0-9_]*$") Then RecordError Errors, ErrorCount, TY_FSPNA, LineN, "Func/Sub/Prop name declared as all lower-case: " & LL
End Sub

Public Sub TestDeclaration(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal L As String, ByVal InSignature As Boolean)
  Dim IsOptional As Boolean, IsByVal As Boolean, IsByRef As Boolean, IsParamArray As Boolean
  Dim IsWithEvents As Boolean, IsEvent As Boolean
  L = Trim(L)
  L = StripLeft(L, "Dim ")
  L = StripLeft(L, "Private ")
  L = StripLeft(L, "Public ")
  L = StripLeft(L, "Const ")
  L = StripLeft(L, "Global ")
  
  Dim Item As Variant, LL As String
  For Each Item In Split(L, ", ")
    Dim IX As Long, ArgName As String, ArgType As String, ArgDefault As String, StandardEvent As Boolean
    LL = Item
    
    IsEvent = RecordLeft(LL, "Event ")
    IsWithEvents = RecordLeft(LL, "WithEvents ")
    IsOptional = RecordLeft(LL, "Optional ")
    IsByVal = RecordLeft(LL, "ByVal ")
    IsByRef = RecordLeft(LL, "ByRef ")
    IsParamArray = RecordLeft(LL, "ParamArray ")
    
    IX = InStr(LL, " = ")
    If IX > 0 Then
      ArgDefault = Trim(Mid(LL, IX + 3))
      LL = Left(LL, IX - 1)
    Else
      ArgDefault = ""
    End If
    
    IX = InStr(LL, " As ")
    If IX > 0 Then
      ArgType = Trim(Mid(LL, IX + 4))
      LL = Left(LL, IX - 1)
    Else
      ArgType = ""
    End If
    
    ArgName = LL
    StandardEvent = IsStandardEvent(ArgName, ArgType)
    
'    If IsParamArray Then Stop
    If ArgType = "" And Not IsEvent And Not StandardEvent Then
      RecordError Errors, ErrorCount, TY_NOTYP, LineN, "Local Parameter Missing Type: [" & ArgName & "]"
      AddFix TY_NOTYP, "\b" & ArgName & "\b", ArgName & " As Variant"
    End If
    If InSignature Then
      If IsParamArray Then
        If Right(LL, 2) <> "()" Then RecordError Errors, ErrorCount, TY_STYLE, LineN, "ParamArray variable not declared as an Array.  Add '()': " & ArgName
      Else
        If Not IsByVal And Not IsByRef And Not StandardEvent Then
          RecordError Errors, ErrorCount, TY_BYRFV, LineN, "ByVal or ByRef not specified on parameter [" & ArgName & "] -- specify one or the other"
          AddFix TY_BYRFV, "\b" & Item & "\b", "ByRef " & Item
        End If
      End If
      If IsOptional And IsByRef Then
        RecordError Errors, ErrorCount, TY_OPBYR, LineN, "Modifiers 'Optional ByRef' may not migrate well: " & ArgName
      End If
      If IsOptional And ArgDefault = "" Then
        RecordError Errors, ErrorCount, TY_OPDEF, LineN, "Parameter declared OPTIONAL but no default specified. Must specify default: " & ArgName
        AddFix TY_OPDEF, "\b" & Item & "\b", Item & " = " & GetTypeDefault(ArgType)
      End If
    End If
    
    TestArgName Errors, ErrorCount, LineN, LL
    
    If Not StandardEvent Then TestArgType Errors, ErrorCount, LineN, LL, ArgType
  Next
End Sub

Public Function GetTypeDefault(ByVal ArgType As String) As String
  Select Case LCase(ArgType)
    Case "string"
      GetTypeDefault = """"""
    Case "long", "integer", "short", "byte", "date", "decimal", "float", "double", "currency"
      GetTypeDefault = "0"
    Case "boolean"
      GetTypeDefault = "False"
    Case "vbtristate"
      GetTypeDefault = "vbUseDefault"
    Case Else
      GetTypeDefault = "Nothing"
  End Select
End Function

Public Function IsStandardEvent(ByVal ArgName As String, ByVal ArgType As String) As Boolean
  If ArgName = "Cancel" Then IsStandardEvent = True: Exit Function
  If ArgName = "LastRow" Then IsStandardEvent = True: Exit Function
  If ArgName = "LastCol" Then IsStandardEvent = True: Exit Function
  If ArgName = "newCol" Then IsStandardEvent = True: Exit Function
  If ArgName = "newCol" Then IsStandardEvent = True: Exit Function
  If ArgName = "newRow" Then IsStandardEvent = True: Exit Function
  If ArgName = "OldValue" Then IsStandardEvent = True: Exit Function
  If ArgName = "Index" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
  If ArgName = "Offset" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
  If ArgName = "UnloadMode" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
  If ArgName = "KeyCode" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
  If ArgName = "KeyAscii" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
  If ArgName = "Button" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
  If ArgName = "Shift" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
  If ArgName = "X" And ArgType = "Single" Then IsStandardEvent = True: Exit Function
  If ArgName = "Y" And ArgType = "Single" Then IsStandardEvent = True: Exit Function
  If ArgName = "Source" And ArgType = "Control" Then IsStandardEvent = True: Exit Function
  If ArgName = "Item" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
  IsStandardEvent = False
End Function

Public Sub TestArgType(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal Name As String, ByVal Typ As String)
  Dim Expect As String
  
  If Typ = "Integer" Then Expect = "Long"
  If Typ = "Short" Then Expect = "Long"
  If Typ = "Byte" Then Expect = "Long"
  If Typ = "Float" Then Expect = "Double"
  If Typ = "Any" Then Expect = "String"
  
  If Expect <> "" Then
    RecordError Errors, ErrorCount, TY_ARGTY, LineN, "Arg [" & Name & "] is of type [" & Typ & "] -- use " & Expect & " (or disable type linting for file)"
  End If
End Sub

Public Sub TestSignature(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal LL As String)
  If Not RegExTest(LL, "^[ ]*(Private|Public|Friend) ") Then RecordError Errors, ErrorCount, TY_PRIPU, LineN, "Either Private or Public should be specified, but neither was."
  
  Dim L As String, WithReturn As Boolean
  L = LL
  L = StripLeft(L, "Private ")
  L = StripLeft(L, "Public ")
  L = StripLeft(L, "Friend ")
  L = StripLeft(L, "Sub ")
  If StartsWith(L, "Function ") Or StartsWith(L, "Property Get ") Then WithReturn = True
  L = StripLeft(L, "Function ")
  L = StripLeft(L, "Property Get ")
  L = StripLeft(L, "Property Let ")
  L = StripLeft(L, "Property Set ")
  
  Dim IX As Long, Ix2 As Long, Name As String, Args As String, Ret As String
  IX = InStr(L, "(")
  If IX = 0 Then Exit Sub
  Name = Left(L, IX - 1)
  If RegExTest(L, "\) As .*\(\)$") Then
    Ix2 = InStrRev(L, ")", Len(L) - 2)
  Else
    Ix2 = InStrRev(L, ")")
  End If
  Args = Mid(L, IX + 1, Ix2 - IX - 1)
  Ret = Mid(L, Ix2 + 1)
  
  TestSignatureName Errors, ErrorCount, LineN, Name
  If WithReturn And Ret = "" Then RecordError Errors, ErrorCount, TY_FNCRE, LineN, "Function Return Type Not Specified -- Specify Return Type or Variant"
  TestDeclaration Errors, ErrorCount, LineN, Args, True
End Sub

Public Sub TestDefaultControlNames(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal Contents As String)
  Dim vTypes() As Variant, vType As Variant
  Dim Matcher As String, Results As String, N As Long, I As Long
  vTypes = Array("CheckBox", "Command", "Option", "Frame", "Label", "TextBox", "RichTextBox", "RichTextBoxNew", "ComboBox", "ListBox", "Timer", "UpDown", "HScrollBar", "Image", "Picture", "MSFlexGrid", "DBGrid", "Line", "Shape", "DTPicker")
  
  For Each vType In vTypes
    Matcher = "Begin [a-zA-Z0-9]*.[a-zA-Z0-9]* " & vType & "[0-9]*"
    N = RegExCount(Contents, Matcher)
    For I = 0 To N - 1
      Results = RegExNMatch(Contents, Matcher, I)
      RecordError Errors, ErrorCount, TY_DFCTL, 0, "Default control name in use on form: " & Results
    Next
  Next
End Sub

Public Sub TestCodeLine(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal L As String)
  If RegExTest(L, "+ """) Or RegExTest(L, """ +") Then RecordError Errors, ErrorCount, TY_CORRE, LineN, "Possible use of + instead of & on String concatenation"
  If RegExTest(L, " Me[.]") Then RecordError Errors, ErrorCount, TY_CORRE, LineN, "Use of 'Me.*' is not required."
  
  If RegExTest(L, "\.Enabled = [-0-9]") Then RecordError Errors, ErrorCount, TY_CORRE, LineN, "Property [Enabled] Should Be Boolean.  Numeric found."
  If RegExTest(L, "\.Visible = [-0-9]") Then RecordError Errors, ErrorCount, TY_CORRE, LineN, "Property [Visible] Should Be Boolean.  Numeric found."

  If RegExTest(L, " Call ") Then RecordError Errors, ErrorCount, TY_CORRE, LineN, "Remove keyword 'Call'."
  If RegExTest(L, " GoSub ") Or RegExTest(L, " Return$") Then RecordError Errors, ErrorCount, TY_GOSUB, LineN, "Remove uses of 'GoSub' and 'Return'."

  If RegExTest(L, " Stop$") Or RegExTest(L, " Return$") Then RecordError Errors, ErrorCount, TY_CSTOP, LineN, "Code contains STOP statement."
End Sub

Public Sub AddFix(ByVal Typ As String, ByVal Find As String, ByVal Repl As String, Optional ByVal RestOfFile As Boolean = False)
  Dim N As Long
  If InStr(AUTOFIX_LINT_TYPES, Typ) = 0 Then Exit Sub
  
On Error Resume Next
  If RestOfFile Then
    N = UBound(AutofixFindRestOfFile)
    N = N + 1
    ReDim Preserve AutofixFindRestOfFile(1 To N)
    ReDim Preserve AutofixReplRestOfFile(1 To N)
    AutofixFindRestOfFile(N) = Find
    AutofixReplRestOfFile(N) = Repl
  Else
    N = UBound(AutofixFind)
    N = N + 1
    ReDim Preserve AutofixFind(1 To N)
    ReDim Preserve AutofixRepl(1 To N)
    AutofixFind(N) = Find
    AutofixRepl(N) = Repl
  End If
End Sub

Public Function GetFixCount(Optional ByVal RestOfFile As Boolean = False) As Long
On Error Resume Next
  GetFixCount = 0
  GetFixCount = UBound(IIf(RestOfFile, AutofixFindRestOfFile, AutofixFind))
End Function

Public Function PerformAutofix(ByVal Line As String) As String
  Dim I As Long, N As Long
    Dim Find As String, Repl As String
  N = GetFixCount(False)
  If N > 0 Then
    For I = LBound(AutofixFind) To UBound(AutofixFind)
      Find = AutofixFind(I)
      Repl = AutofixRepl(I)
      If Find = "" Then GoTo NextFix
      Line = RegExReplace(Line, Find, Repl)
NextFix:
    Next
  End If
  
  N = GetFixCount(True)
  If N > 0 Then
    For I = LBound(AutofixFindRestOfFile) To UBound(AutofixFindRestOfFile)
      Find = AutofixFindRestOfFile(I)
      Repl = AutofixReplRestOfFile(I)
      If Find = "" Then GoTo NextFixRestOfFile
      Line = RegExReplace(Line, Find, Repl)
NextFixRestOfFile:
    Next
  End If
  
Finish:
  PerformAutofix = Line
  
  Erase AutofixFind
  Erase AutofixRepl
End Function
