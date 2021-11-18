Attribute VB_Name = "modQuickLint"
Option Explicit
 
Private Const Idnt As Long = 2
Private Const MAX_ERRORS As Long = 10
Private Const Attr As String = "Attribute"
Private Const Q As String = """"
Private Const A As String = "'"
Private Const S As String = " "

Private Const TY_ERROR As String = "Error"
Private Const TY_INDNT As String = "Indnt"
Private Const TY_ARGNA As String = "ArgNa"
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
Private Const TY_DFCTL As String = "DfCtl"

Public ErrorPrefix As String

Public Function Lint(Optional ByVal FileName As String = "", Optional ByVal Alert As Boolean = True) As String
  Dim FileList As String
  If FileName = "" Then FileName = "prj.vbp"
  If InStr(FileName, "\") = 0 Then FileName = App.Path & "\" & FileName
  FileList = IIf(Right(FileName, 4) = ".vbp", VBPCode(FileName), FileName)
  
  Lint = QuickLintFiles(FileList)
End Function

Public Function QuickLintFiles(ByVal List As String) As String
  Const lintDotsPerRow As Long = 50
  Dim L As Variant
  Dim X As Long
  Dim StartTime As Date
  StartTime = Now
  
  For Each L In Split(List, vbCrLf)
    Dim Result As String
    Result = QuickLintFile(L)
    If Not Result = "" Then
      Dim S As String
      Debug.Print vbCrLf & "Done (" & DateDiff("s", StartTime, Now) & "s).  To re-run for failing file, hit enter on the line below:"
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

Public Function QuickLintFile(ByVal File As String) As String
  If InStr(File, "\") = 0 Then File = App.Path & "\" & File
  Dim FName As String, Contents As String, GivenName As String, CheckName As String
  FName = Mid(File, InStrRev(File, "\") + 1)
  CheckName = Replace(Replace(Replace(FName, ".bas", ""), ".cls", ""), ".frm", "")
  ErrorPrefix = Right(Space(15) & FName, 15) & " "
  Contents = ReadEntireFile(File)
  GivenName = RegExNMatch(Contents, "Attribute VB_Name = ""([^""]+)""", 0)
  GivenName = Replace(Replace(GivenName, "Attribute VB_Name = ", ""), """", "")
  If CheckName <> GivenName Then
    QuickLintFile = "Module name [" & GivenName & "] must match file name [" & FName & "].  Rename module or class to match the other"
    Exit Function
  End If
  QuickLintFile = QuickLintContents(Contents)
End Function

Public Function QuickLintContents(ByVal Contents As String) As String
  Dim Lines() As String, LL As Variant, L As String
On Error GoTo LintError
  Lines = Split(Replace(Contents, vbCr, ""), vbLf)
  
  Dim InAttributes As Boolean, InBody As Boolean
    
  Dim MultiLine As String
  Dim Indent As Long, LineN As Long
  Dim Errors As String, ErrorCount As Long
  Dim BlankLineCount As Long
  Dim Options As New Collection
  Indent = 0
  
  TestDefaultControlNames Errors, ErrorCount, 0, Contents

  
  For Each LL In Lines
    If ErrorCount >= MAX_ERRORS Then Exit For
    
    If Right(LL, 2) = " _" Then
      Dim Portion As String
      Portion = Left(LL, Len(LL) - 2)
      If MultiLine <> "" Then Portion = Trim(Portion)
      MultiLine = MultiLine + Portion
      LineN = LineN + 1
      GoTo NextLine
    ElseIf MultiLine <> "" Then
      LL = MultiLine & LL
      MultiLine = ""
    End If
    
    TestBlankLines Errors, ErrorCount, LineN, LL, BlankLineCount
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
    'If LineN = 15 Then Stop
    
    Dim UnindentedAlready As Boolean
    If RegExTest(L, "^Option ") Then
      Options.Add "true", Replace(L, "Options ", "")
    ElseIf RegExTest(L, "^[ ]*(Else|ElseIf .* Then)$") Then
      Indent = Indent - Idnt
    ElseIf RegExTest(L, "^[ ]*End Select$") Then
      Indent = Indent - Idnt - Idnt
    ElseIf RegExTest(L, "^[ ]*(End (If|Function|Sub|Property|Enum|Type)|Next( .*)?|Wend|Loop|Loop (While .*|Until .*))$") Then
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
      
      If RegExTest(L, "^[ ]*(Else|ElseIf .* Then)$") Then
        Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*(End (If|Function|Sub|Property)|Next|Wend|Loop|Loop .*|Enum|Type|Select)$") Then
        If Not UnindentedAlready Then Indent = Indent - Idnt
      ElseIf RegExTest(St, "^[ ]*If ") Then
        If Not RegExTest(St, "Then ") Then Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*For ") Then
        If Not RegExTest(St, " Next") Then Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*Next$") Then
        Indent = Indent - Idnt
      ElseIf RegExTest(St, "^[ ]*Next [a-zA-Z_][a-zA-Z0-9_]*$") Then
        RecordError Errors, ErrorCount, TY_STYLE, LineN, "Remove variable from NEXT statement"
        Indent = Indent - Idnt
      ElseIf RegExTest(St, "^[ ]*While ") Then
        RecordError Errors, ErrorCount, TY_STYLE, LineN, "Use Do While/Until...Loop in place of While...Wend"
        If Not RegExTest(St, " Wend$") Then Indent = Indent + Idnt
      ElseIf RegExTest(St, "^[ ]*Do (While|Until)") Then
        If Not RegExTest(St, ": Loop") Then Indent = Indent + Idnt
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
  Next
  
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
    
    Line = Left(Line, X - 1) & String(Y - X + 1, "S") & Mid(Line, Y + 1)
  Loop
  
  X = InStr(Line, A)
  If X > 0 Then Line = RTrim(Left(Line, X - 1))
  
  CleanLine = Line
End Function
  
Public Sub RecordError(ByRef Errors As String, ByRef ErrorCount As Long, ByVal Typ As String, ByVal LineN As Long, ByVal Error As String)
  If Len(Errors) <> 0 Then Errors = Errors & vbCrLf
  Errors = Errors & ErrorPrefix & "[" & Right(Space(5) & Typ, 5) & "] Line " & Right(Space(5) & LineN, 5) & ": " & Error
  ErrorCount = ErrorCount + 1
End Sub

Public Function StartsWith(ByVal L As String, ByVal Find As String) As Boolean
  StartsWith = Left(L, Len(Find)) = Find
End Function

Public Function StripLeft(ByVal L As String, ByVal Find As String) As String
  If StartsWith(L, Find) Then StripLeft = Mid(L, Len(Find) + 1) Else StripLeft = L
End Function

Public Sub TestIndent(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal L As String, ByVal LineIndent As Long, ByVal ExpectedIndent As Long)
  If RTrim(L) = "" Then Exit Sub
  If RegExTest(L, "^On Error ") Then Exit Sub
  If RegExTest(L, "^[a-zA-Z][a-zA-Z0-9]*:$") Then Exit Sub
    
  If LineIndent <> ExpectedIndent Then RecordError Errors, ErrorCount, TY_INDNT, LineN, "Incorrect Indent -- expected " & ExpectedIndent & ", got " & LineIndent
End Sub

Public Sub TestBlankLines(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal L As String, ByRef BlankLineCount As Long)
  If Trim(L) <> "" Then
    BlankLineCount = 0
    Exit Sub
  End If
  BlankLineCount = BlankLineCount + 1
  If BlankLineCount > 3 Then RecordError Errors, ErrorCount, TY_BLANK, LineN, "Too many blank lines."
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
  
  If RegExTest(LL, "^[a-z][a-z0-9_]*$") Then RecordError Errors, ErrorCount, TY_ARGNA, LineN, "Identifier name declared as all lower-case"
  
  If RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*%$") Then ' % Integer Dim L%
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Integer deprecated: " & LL
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*&$") Then ' & Long  Dim M&
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Long deprecated: " & LL
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*@$") Then ' @ Decimal Const W@ = 37.5
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Decimal deprecated: " & LL
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-TY_TYPEC-Z0-9_]*!$") Then ' ! Single  Dim Q!
    RecordError Errors, ErrorCount, TY_DEPRE, LineN, "Use of Type Character For Single deprecated: " & LL
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*#$") Then ' # Double  Dim X#
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Double deprecated: " & LL
  ElseIf RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*\$$") Then ' $ String  Dim V$ = "Secret"
    RecordError Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For String deprecated: " & LL
  End If
End Sub

Public Sub TestSignatureName(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal Name As String)
  Dim LL As String
  LL = Trim(Name)
  
  If RegExTest(LL, "^[a-z][a-z0-9_]*$") Then RecordError Errors, ErrorCount, TY_FSPNA, LineN, "Func/Sub/Prop name declared as all lower-case: " & Name
End Sub

Public Sub TestDeclaration(ByRef Errors As String, ByRef ErrorCount As Long, ByVal LineN As Long, ByVal L As String, ByVal InSignature As Boolean)
  Dim IsOptional As Boolean, IsByVal As Boolean, IsByRef As Boolean, IsParamArray As Boolean
  L = Trim(L)
  L = StripLeft(L, "Dim ")
  L = StripLeft(L, "Private ")
  L = StripLeft(L, "Public ")
  L = StripLeft(L, "Const ")
  L = StripLeft(L, "Global ")
  
  Dim LL As Variant
  For Each LL In Split(L, ", ")
    Dim Ix As Long, ArgName As String, ArgType As String, ArgDefault As String
    
    IsOptional = StartsWith(LL, "Optional ")
    LL = StripLeft(LL, "Optional ")
    
    IsByVal = StartsWith(LL, "ByVal ")
    LL = StripLeft(LL, "ByVal ")
    
    IsByRef = StartsWith(LL, "ByRef ")
    LL = StripLeft(LL, "ByRef ")
    
    IsParamArray = StartsWith(LL, "ParamArray ")
    LL = StripLeft(LL, "ParamArray ")
    
    Ix = InStr(LL, " = ")
    If Ix > 0 Then
      ArgDefault = Trim(Mid(LL, Ix + 3))
      LL = Left(LL, Ix - 1)
    Else
      ArgDefault = ""
    End If
    
    Ix = InStr(LL, " As ")
    If Ix > 0 Then
      ArgType = Trim(Mid(LL, Ix + 4))
      LL = Left(LL, Ix - 1)
    Else
      ArgType = ""
    End If
    
'    If IsParamArray Then Stop
    If ArgType = "" Then RecordError Errors, ErrorCount, TY_NOTYP, LineN, "Local Parameter Missing Type: [" & LL & "]"
    If InSignature Then
      If IsParamArray Then
        If Right(LL, 2) <> "()" Then RecordError Errors, ErrorCount, TY_STYLE, LineN, "ParamArray variable not declared as an Array.  Add '()': " & LL
      Else
        If Not IsByVal And Not IsByRef Then RecordError Errors, ErrorCount, TY_BYRFV, LineN, "ByVal or ByRef not specified on praameter [" & LL & "] -- specify one or the other"
      End If
      If IsOptional And ArgDefault = "" Then RecordError Errors, ErrorCount, TY_OPDEF, LineN, "Parameter declared OPTIONAL but no default specified. Must specify default: " & LL
    End If
    
    TestArgName Errors, ErrorCount, LineN, LL
  Next
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
  L = StripLeft(L, "Property ")
  
  Dim Ix As Long, Ix2 As Long, Name As String, Args As String, Ret As String
  Ix = InStr(L, "(")
  If Ix = 0 Then Exit Sub
  Name = Left(L, Ix - 1)
  If RegExTest(L, "\) As .*\(\)") Then
    Ix2 = InStrRev(L, ")", Len(L) - 2)
  Else
    Ix2 = InStrRev(L, ")")
  End If
  Args = Mid(L, Ix + 1, Ix2 - Ix - 1)
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
  If RegExTest(L, "+ """) Or RegExTest(L, """ +") Then RecordError Error, ErrorCount, TY_CORRE, LineN, "Possible use of + instead of & on String concatenation"
  If RegExTest(L, " Me[.]") Then RecordError Error, ErrorCount, TY_CORRE, LineN, "Use of 'Me.*' is not required."
  
  If RegExTest(L, "\.Enabled = [-0-9]") Then RecordError Error, ErrorCount, TY_CORRE, LineN, "Property [Enabled] Should Be Boolean.  Numeric found."
  If RegExTest(L, "\.Visible = [-0-9]") Then RecordError Error, ErrorCount, TY_CORRE, LineN, "Property [Visible] Should Be Boolean.  Numeric found."

  If RegExTest(L, " Call ") Then RecordError Error, ErrorCount, TY_CORRE, LineN, "Remove keyword 'Call'."
  If RegExTest(L, " GoSub ") Or RegExTest(L, " Return$") Then RecordError Error, ErrorCount, TY_GOSUB, LineN, "Remove uses of 'GoSub' and 'Return'."

  If RegExTest(L, " Stop$") Or RegExTest(L, " Return$") Then RecordError Error, ErrorCount, TY_CSTOP, LineN, "Code contains STOP statement."
End Sub

