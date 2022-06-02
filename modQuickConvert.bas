Attribute VB_Name = "modQuickConvert"
Option Explicit
 
Private Const Idnt As Long = 2
Private Const Attr As String = "Attribute"
Private Const Q As String = """"
Private Const A As String = "'"
Private Const S As String = " "

Private Const STRING_TOKEN_PREFIX As String = "__S"
Private Const EXPRESSION_TOKEN_PREFIX As String = "__E_"

Private LineStrings() As String, LineStringsCount As Long
Private LineComment As String
Private InProperty As Boolean
Private CurrentTypeName As String
Private CurrentEnumName As String
Private CurrentFunctionName As String
Private CurrentFunctionReturnValue As String
Private CurrentFunctionArgs As String
Private CurrentFunctionArrays As String
Private ModuleName As String
Private ModuleFunctions As String
Private ModuleArrays As String
Private ModuleProperties As String
Private WithVars As String

Public Enum DeclarationType
  DECL_GLOBAL = 99
  DECL_SIGNATURE = 98
  DECL_LOCAL = 1
  DECL_TYPE
  DECL_ENUM
  DECL_EXTERN = 101
End Enum

Public Enum CodeType
  CODE_MODULE
  CODE_CLASS
  CODE_FORM
  CODE_CONTROL
End Enum

Public Type RandomType
  J As Long
  W As String
  X As String * 5
End Type

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

Public Function Convert(Optional ByVal FileName As String = "") As String
  Dim FileList As String
  FileList = ResolveSources(FileName)
  Convert = QuickConvertFiles(FileList)
End Function

Public Function QuickConvertFiles(ByVal List As String) As String
  Const lintDotsPerRow As Long = 50
  Dim L As Variant
  Dim X As Long
  Dim StartTime As Date
  StartTime = Now
  
  For Each L In Split(List, vbCrLf)
    Dim Result As String
    Result = QuickConvertFile(L)
    If Result <> "" Then
      Dim S As String
      Debug.Print vbCrLf & "Done (" & DateDiff("s", StartTime, Now) & "s).  To re-run for failing file, hit enter on the line below:"
      S = "LINT FAILED: " & L & vbCrLf & Result & vbCrLf & "?Lint(""" & L & """)"
      QuickConvertFiles = S
      Exit Function
    Else
      Debug.Print Switch(Right(L, 3) = "frm", "o", Right(L, 3) = "cls", "x", Right(L, 3) = "ctl", "+", True, ".");
    End If
    X = X + 1
    If X >= lintDotsPerRow Then X = 0: Debug.Print
    DoEvents
  Next
  Debug.Print vbCrLf & "Done (" & DateDiff("s", StartTime, Now) & "s)."
  QuickConvertFiles = ""
End Function

Public Function CodeFileType(ByVal File As String) As CodeType
  Select Case Right(LCase(File), 4)
    Case ".bas": CodeFileType = CODE_MODULE
    Case ".frm": CodeFileType = CODE_FORM
    Case ".cls": CodeFileType = CODE_CLASS
    Case ".ctl": CodeFileType = CODE_CONTROL
    Case Else: CodeFileType = CODE_MODULE
  End Select
End Function

Public Function QuickConvertFile(ByVal File As String) As String
  ModuleArrays = ""

  If InStr(File, "\") = 0 Then File = App.Path & "\" & File
  Dim fName As String, Contents As String, GivenName As String, CheckName As String
  fName = Mid(File, InStrRev(File, "\") + 1)
  CheckName = Replace(Replace(Replace(fName, ".bas", ""), ".cls", ""), ".frm", "")
  ErrorPrefix = Right(Space(18) & fName, 18) & " "
  Contents = ReadEntireFile(File)
  GivenName = GetModuleName(Contents)
  If LCase(CheckName) <> LCase(GivenName) Then
    QuickConvertFile = "QuickConvertFile: Module name [" & GivenName & "] must match file name [" & fName & "].  Rename module or class to match the other"
    Exit Function
  End If
  QuickConvertFile = ConvertContents(Contents, CodeFileType(File))
End Function

Public Function GetModuleName(ByVal Contents As String) As String
  GetModuleName = RegExNMatch(Contents, "Attribute VB_Name = ""([^""]+)""", 0)
  GetModuleName = Replace(Replace(GetModuleName, "Attribute VB_Name = ", ""), """", "")
End Function

Public Function I(ByVal N As Long) As String
  If N <= 0 Then I = "" Else I = Space(N)
End Function

Public Function ConvertContents(ByVal Contents As String, ByVal vCodeType As CodeType, Optional ByVal SubSegment As Boolean = False) As String
  Dim Lines() As String, ActualLine As Variant, LL As String, L As String
'On Error GoTo LintError
  If Not SubSegment Then
    ModuleName = GetModuleName(Contents)
    ModuleFunctions = GetModuleFunctions(Contents)
  End If
  
  Lines = Split(Replace(Contents, vbCr, ""), vbLf)

  Dim InAttributes As Boolean, InBody As Boolean
  InBody = SubSegment

  Dim MultiLineOrig As String, MultiLine As String, IsMultiLine As Boolean
  Dim LineN As Long, Indent As Long
  Dim NewContents As String
  Dim SelectHasCase As Boolean
  
  Indent = 0
  NewContents = ""
'  NewContents = UsingEverything & vbCrLf2
'  NewContents = NewContents & "static class " & ModuleName & " {" & vbCrLf

  For Each ActualLine In Lines
    LL = ActualLine
'    If MaxErrors > 0 And ErrorCount >= MaxErrors Then Exit For
    
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
    
    L = CleanLine(LL)
    
    If Not InBody Then
      Dim IsAttribute As Boolean
      IsAttribute = StartsWith(LTrim(L), "Attribute ")
      If Not InAttributes And IsAttribute Then
        InAttributes = True
        GoTo NextLineWithoutRecord
      ElseIf InAttributes And Not IsAttribute Then
        InAttributes = False
        InBody = True
        LineN = 0
      Else
        GoTo NextLineWithoutRecord
      End If
    End If
    
    LineN = LineN + 1
'    If LineN >= 357 Then Stop
    
    Dim UnindentedAlready As Boolean
    
    If RegExTest(L, "^[ ]*(Else|ElseIf .* Then)$") Then
      Indent = Indent - Idnt
      UnindentedAlready = True
    ElseIf RegExTest(L, "^[ ]*End Select$") Then
      Indent = Indent - Idnt - Idnt
    ElseIf RegExTest(L, "^[ ]*(End (If|Function|Sub|Property|Enum|Type|With)|Next( .*)?|Wend|Loop|Loop (While .*|Until .*)|ElseIf .*)$") Then
      Indent = Indent - Idnt
      UnindentedAlready = True
      CurrentEnumName = ""
      CurrentTypeName = ""
      If RegExTest(L, "^[ ]*End With") Then Stack WithVars
    Else
      UnindentedAlready = False
    End If
    
    Dim NewLine As String
    NewLine = ""
    
    If InProperty Then ' we process properties out of band to keep getters and setters together
      If InStr(L, "End Property") > 0 Then InProperty = False
      GoTo NextLineWithoutRecord
    End If
    
    If CurrentTypeName <> "" Then ' if we are in a type or an enum, the entire line is parsed as such
      NewLine = NewLine & ConvertTypeLine(L, vCodeType)
    ElseIf CurrentEnumName <> "" Then
      NewLine = NewLine & ConvertEnumLine(L)
    ElseIf RegExTest(L, "^[ ]*If ") Then ' The "If" control structure, when single-line, lacks the "End If" to signal a close.
      NewLine = NewLine & ConvertIf(L)
      If InStr(L, " Then ") = 0 Then Indent = Indent + Idnt
    ElseIf RegExTest(L, "^[ ]*ElseIf .*$") Then
      NewLine = NewLine & ConvertIf(L)
      If InStr(L, " Then ") = 0 Then Indent = Indent + Idnt
    Else
      Dim Statements() As String, SSI As Long, St As String
      Statements = Split(Trim(L), ": ")
      For SSI = LBound(Statements) To UBound(Statements)
        St = Statements(SSI)
        
        If RegExTest(St, "^[ ]*ElseIf .*$") Then
          NewLine = NewLine & ConvertIf(St)
          Indent = Indent + Idnt
        ElseIf RegExTest(St, "^[ ]*Else$") Then
          NewLine = NewLine & "} else {"
          Indent = Indent + Idnt
        ElseIf RegExTest(St, "^[ ]*End Function") Then
          NewLine = NewLine & "return " & CurrentFunctionReturnValue & ";" & vbCrLf & "}"
          CurrentFunctionName = ""
          CurrentFunctionReturnValue = ""
          CurrentFunctionArrays = ""
          If Not UnindentedAlready Then Indent = Indent - Idnt
        ElseIf RegExTest(St, "^[ ]*End Select$") Then
          NewLine = NewLine & "break;" & vbCrLf
          NewLine = NewLine & "}"
          If Not UnindentedAlready Then Indent = Indent - Idnt
        ElseIf RegExTest(St, "^[ ]*End (If|Sub|Enum|Type)$") Then
          CurrentTypeName = ""
          CurrentEnumName = ""
          NewLine = NewLine & "}"
          If Not UnindentedAlready Then Indent = Indent - Idnt
        ElseIf RegExTest(St, "^[ ]*For Each") Then
          Indent = Indent + Idnt
          NewLine = ConvertForEach(St)
        ElseIf RegExTest(St, "^[ ]*For ") Then
          Indent = Indent + Idnt
          NewLine = ConvertFor(St)
        ElseIf RegExTest(St, "^[ ]*Next\b") Then
          NewLine = NewLine & "}"
          If Not UnindentedAlready Then Indent = Indent - Idnt
        ElseIf RegExTest(St, "^[ ]*While ") Then
          NewLine = NewLine & ConvertWhile(St)
          Indent = Indent + Idnt
        ElseIf RegExTest(St, "^[ ]*Wend") Then
          NewLine = NewLine & "}"
          If Not UnindentedAlready Then Indent = Indent - Idnt
        ElseIf RegExTest(St, "^[ ]*Do (While|Until)") Then
          NewLine = NewLine & ConvertWhile(St)
          Indent = Indent + Idnt
        ElseIf RegExTest(St, "^[ ]*Loop$") Then
          NewLine = NewLine & "}"
        ElseIf RegExTest(St, "^[ ]*Do$") Then
          NewLine = NewLine & "do {"
          Indent = Indent + Idnt
        ElseIf RegExTest(St, "^[ ]*(Loop While |Loop Until )") Then
          NewLine = NewLine & ConvertWhile(St)
        ElseIf RegExTest(St, "^[ ]*With ") Then
          NewLine = NewLine & ConvertWith(St)
          Indent = Indent + Idnt
        ElseIf RegExTest(St, "^[ ]*Select Case ") Then
          NewLine = NewLine & ConvertSwitch(St)
          Indent = Indent + Idnt + Idnt
          SelectHasCase = False
        ElseIf RegExTest(St, "^[ ]*Case ") Then
          NewLine = NewLine & ConvertSwitchCase(St, SelectHasCase)
          SelectHasCase = True
        ElseIf RegExTest(St, "^[ ]*(Private |Public )?Declare (Function |Sub )") Then
          NewLine = NewLine & ConvertDeclare(St) ' External Api
        ElseIf RegExTest(St, "^((Private|Public|Friend) )?Function ") Then
          CurrentFunctionArgs = ""
          Indent = Indent + Idnt
          NewLine = NewLine & ConvertSignature(St, vCodeType)
        ElseIf RegExTest(St, "^((Private|Public|Friend) )?Sub ") Then
          CurrentFunctionArgs = ""
          Indent = Indent + Idnt
          NewLine = NewLine & ConvertSignature(St, vCodeType)
        ElseIf RegExTest(St, "^((Private|Public|Friend) )?Property (Get|Let|Set) ") Then
          CurrentFunctionArgs = ""
          NewLine = NewLine & ConvertProperty(St, Contents, vCodeType)
          InProperty = Not EndsWith(L, "End Property")
          If InProperty Then
            Indent = Indent + Idnt
          Else
            GoTo NextLine
          End If
        ElseIf RegExTest(St, "^[ ]*(Public |Private )?Enum ") Then
          NewLine = NewLine & ConvertEnum(St)
          Indent = Indent + Idnt
        ElseIf RegExTest(St, "^[ ]*(Public |Private )?Type ") Then
          NewLine = NewLine & ConvertType(St)
          Indent = Indent + Idnt
        ElseIf RegExTest(St, "^[ ]*(Dim|Private|Public|Const|Global|Static) ") Then
          NewLine = NewLine & ConvertDeclaration(St, IIf(CurrentFunctionName = "", DECL_GLOBAL, DECL_LOCAL), vCodeType)
        Else
          NewLine = NewLine & ConvertStatement(St)
        End If
NextStatement:
      Next
    End If
NextLine:
'      If IsMultiLine Then Stop
'      If InStr(LL, "Function") > 0 Then Stop
'      If InStr(LL, "Private Function") > 0 Then Stop
'    If Indent < 0 Then Stop
    NewLine = Decorate(NewLine)
    If Trim(NewLine) <> "" Then
      NewContents = NewContents & I(Indent) & NewLine & vbCrLf
    End If
NextLineWithoutRecord:
  Next
'  If AutoFix <> "" Then WriteFile AutoFix, Left(NewContents, Len(NewContents) - 2), True
  
'  NewContents = NewContents & "}" & vbCrLf
  ConvertContents = NewContents
  Exit Function
LintError:
  Debug.Print "Error in quick convert [" & Err.Number & "]: " & Err.Description
  ConvertContents = "Error in quick convert [" & Err.Number & "]: " & Err.Description
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

' de string and decomment a given line (before conversion)
Public Function CleanLine(ByVal Line As String) As String
  Dim X As Long, Y As Long, Token As String, Value As String
  
  Erase LineStrings
  LineStringsCount = 0
  LineComment = ""
  
  Do While True
    X = InStr(Line, Q)
    If X = 0 Then Exit Do
    
    Y = InStr(X + 1, Line, Q)
    Do While Mid(Line, Y + 1, 1) = Q
      Y = InStr(Y + 2, Line, Q)
    Loop
    
    If Y = 0 Then Exit Do
    
    LineStringsCount = LineStringsCount + 1
    ReDim Preserve LineStrings(1 To LineStringsCount)
    Value = ConvertStringLiteral(Mid(Line, X, Y - X + 1))
    LineStrings(LineStringsCount) = Value
    Token = STRING_TOKEN_PREFIX & LineStringsCount
    Line = Left(Line, X - 1) & Token & Mid(Line, Y + 1)
  Loop
  
  X = InStr(Line, A)
  If X > 0 Then
    LineComment = Trim(Mid(Line, X + 1))
    Line = RTrim(Left(Line, X - 1))
  End If
  
  CleanLine = Line
End Function

' re-string and re-comment a given line (after conversion)
Public Function Decorate(ByVal Line As String) As String
  Dim I As Long
  For I = LineStringsCount To 1 Step -1
    Line = Replace(Line, "__S" & I, LineStrings(I))
  Next
  
  If LineComment <> "" Then Line = Line & " // " & LineComment
  Decorate = Line
End Function

Public Function ConvertStringLiteral(ByVal L As String) As String
  L = Replace(L, "\", "\\")
  L = """" & Replace(Mid(L, 2, Len(L) - 2), """""", "\""") & """"
  ConvertStringLiteral = L
End Function
  
Public Function StartsWith(ByVal L As String, ByVal Find As String) As Boolean
  StartsWith = Left(L, Len(Find)) = Find
End Function

Public Function EndsWith(ByVal L As String, ByVal Find As String) As Boolean
  EndsWith = Right(L, Len(Find)) = Find
End Function

Public Function StripLeft(ByVal L As String, ByVal Find As String) As String
  If StartsWith(L, Find) Then StripLeft = Mid(L, Len(Find) + 1) Else StripLeft = L
End Function

Public Function RecordLeft(ByRef L As String, ByVal Find As String) As Boolean
  RecordLeft = StartsWith(L, Find)
  If RecordLeft Then L = Mid(L, Len(Find) + 1)
End Function

Public Function RemoveUntil(ByRef L As String, ByVal Find As String, Optional ByVal RemoveFind As Boolean = False) As String
  Dim IX As Long
  IX = InStr(L, Find)
  If IX <= 0 Then Exit Function
  RemoveUntil = Left(L, IX - 1)
  L = Mid(L, IIf(RemoveFind, IX + Len(Find), IX))
End Function

Private Function GetModuleFunctions(ByVal Contents As String) As String
  Const Pattern As String = "(Private (Function|Sub) [^(]+\()"
  Dim N As Long, I As Long
  Dim S As String
  N = RegExCount(Contents, Pattern)
  GetModuleFunctions = ""
  For I = 0 To N - 1
    S = RegExNMatch(Contents, Pattern, I)
    S = Replace(S, "Private ", "")
    S = Replace(S, "Sub ", "")
    S = Replace(S, "Function ", "")
    S = Replace(S, "(", "")
    GetModuleFunctions = GetModuleFunctions & "[" & S & "]"
  Next
End Function

Private Function IsLocalFuncRef(ByVal F As String) As Boolean
  IsLocalFuncRef = InStr(ModuleFunctions, "[" & Trim(F) & "]") <> 0
End Function

Private Function SearchLeft(ByVal Start As Long, ByVal Src As String, ByVal Find As String, Optional ByVal NotIn As Boolean = False, Optional ByVal Reverse As Boolean = False) As Long
  Dim Bg As Long, Ed As Long, St As Long
  Dim I As Long, C As String, Found As Boolean
  If Not Reverse Then
    Bg = IIf(Start = 0, 1, Start)
    Ed = Len(Src)
    St = 1
  Else
    Bg = IIf(Start = 0, Len(Src), Start)
    Ed = 1
    St = -1
  End If
  
  For I = Bg To Ed Step St
    C = Mid(Src, I, 1)
    Found = InStr(Find, C) > 0
    If Not NotIn And Found Or NotIn And Not Found Then
      SearchLeft = I
      Exit Function
    End If
  Next
  
  SearchLeft = 0
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function ConvertIf(ByVal L As String) As String
  Dim ixThen As Long, Expression As String
  Dim WithThen As Boolean, WithElse As Boolean
  Dim MultiStatement As Boolean
  L = Trim(L)
  ixThen = InStr(L, " Then")
  WithThen = InStr(L, " Then ") > 0
  WithElse = InStr(L, " Else ") > 0
  Expression = Trim(Left(L, ixThen - 1))
  Expression = StripLeft(Expression, "If ")
  Expression = StripLeft(Expression, "ElseIf ")
  
  ConvertIf = IIf(Not IsInStr(L, "ElseIf"), "if", "} else if")
  ConvertIf = ConvertIf & "(" & ConvertExpression(Expression) & ")"
  
  If Not WithThen Then
    ConvertIf = ConvertIf & " {"
  Else
    Dim cThen As String, cElse As String
    cThen = Trim(Mid(L, ixThen + 5))
    Dim ixElse As Long
    ixElse = InStr(cThen, " Else ")
    If ixElse > 0 Then
      cElse = Mid(cThen, ixElse + 6)
      cThen = Left(cThen, ixElse - 1)
    Else
      cElse = ""
    End If
      
    ' Inline Then
    Dim St As Variant
    MultiStatement = InStr(cThen, ": ") > 0
    If MultiStatement Then
      ConvertIf = ConvertIf & " { "
      For Each St In Split(cThen, ": ")
        ConvertIf = ConvertIf & ConvertStatement(St) & " "
      Next
      ConvertIf = ConvertIf & "}"
    Else
      ConvertIf = ConvertIf & ConvertStatement(cThen)
    End If
    
    ' Inline Then ... Else
    If ixElse > 0 Then
      MultiStatement = InStr(cElse, ":") > 0
      If MultiStatement Then
        ConvertIf = ConvertIf & " else { "
        For Each St In Split(cElse, ":")
          ConvertIf = ConvertIf & ConvertStatement(Trim(St))
        Next
        ConvertIf = ConvertIf & " }"
      Else
        ConvertIf = ConvertIf & " else " & ConvertStatement(cElse)
      End If
    End If
  End If
  
End Function

Public Function ConvertWith(ByVal L As String) As String
  Dim Value As String
  Value = Trim(L)
  Value = StripLeft(Value, "With ")

  If ValueIsSimple(Value) Then
    WithVars = Stack(WithVars, ConvertExpression(Value))
    ConvertWith = "// Converted WITH statement: " & L
  Else
    Dim WithVar As String
    WithVar = "__withVar" & Random(1000)
    ConvertWith = ""
    ConvertWith = ConvertWith & "// " & L & " // TODO (not supported): Expression used in WITH.  Verify result: " + Value
    ConvertWith = ConvertWith & vbCrLf & "dynamic " & WithVar & " = " & ConvertStatement(Value) & ";"
    WithVars = Stack(WithVars, WithVar)
  End If
End Function

Public Function WithVar() As String
  WithVar = Stack(WithVars, , True)
End Function

Public Function ConvertSwitch(ByVal L As String) As String
  ConvertSwitch = "switch(" & ConvertExpression(Trim(Replace(L, "Select Case ", ""))) & ") {"
End Function

Public Function ConvertSwitchCase(ByVal L As String, ByVal SelectHasCase As Boolean) As String
  Dim V As Variant
  ConvertSwitchCase = ""
  If SelectHasCase Then ConvertSwitchCase = ConvertSwitchCase & "break;" & vbCrLf
  If Trim(L) = "Case Else" Then
    ConvertSwitchCase = ConvertSwitchCase & "default: "
  Else
    RecordLeft L, "Case "
    If Right(L, 1) = ":" Then L = Left(L, Len(L) - 1)
    For Each V In Split(L, ", ")
      V = Trim(V)
      If InStr(V, " To ") > 0 Then
        ConvertSwitchCase = ConvertSwitchCase & "default: /* TODO: Cannot Convert Ranged Case: " & L & " */"
      ElseIf StartsWith(V, "Is ") Then
        ConvertSwitchCase = ConvertSwitchCase & "default: /* TODO: Cannot Convert Expression Case: " & L & " */"
      Else
        ConvertSwitchCase = ConvertSwitchCase & "case " & ConvertExpression(V) & ": "
      End If
    Next
  End If
End Function

Public Function ConvertWhile(ByVal L As String) As String
  Dim Exp As String, Closing As Boolean, Invert As Boolean
  L = LTrim(L)
  If RecordLeft(L, "Do While ") Then
    Exp = L
  ElseIf RecordLeft(L, "Do Until ") Then
    Exp = L
    Invert = True
  ElseIf RecordLeft(L, "While ") Then
    Exp = L
  ElseIf RecordLeft(L, "Loop While ") Then
    Exp = L
    Closing = True
  ElseIf RecordLeft(L, "Loop Until ") Then
    Exp = L
    Closing = True
    Invert = True
  End If
  
  ConvertWhile = ""
  If Closing Then ConvertWhile = ConvertWhile & "} "
  ConvertWhile = ConvertWhile & "while("
  If Invert Then ConvertWhile = ConvertWhile & "!("
  ConvertWhile = ConvertWhile & ConvertExpression(Exp)
  If Invert Then ConvertWhile = ConvertWhile & ")"
  ConvertWhile = ConvertWhile & ")"
  If Not Closing Then ConvertWhile = ConvertWhile & " {" Else ConvertWhile = ConvertWhile & ";"
End Function

Public Function ConvertFor(ByVal L As String) As String
  Dim Var As String, ForFrom As String, ForTo As String, ForStep As String
  Dim ForReverse As Boolean, ForCheck As String
  L = Trim(L)
  RecordLeft L, "For "
  
  Var = RemoveUntil(L, " = ", True)
  ForFrom = RemoveUntil(L, " To ", True)
  ForTo = L
  
  ForStep = RemoveUntil(ForTo, " Step ", True)
  If ForStep = "" Then ForStep = "1"
  
  ForStep = ConvertExpression(ForStep)
  ForReverse = InStr(ForStep, "-") > 0
  If ForReverse Then ForCheck = " >= " Else ForCheck = " <= "
  
  ConvertFor = ""
  ConvertFor = ConvertFor & "for ("
  ConvertFor = ConvertFor & ExpandToken(Var) & " = " & ConvertExpression(ForFrom) & "; "
  ConvertFor = ConvertFor & ExpandToken(Var) & ForCheck & ConvertExpression(ForTo) & "; "
  ConvertFor = ConvertFor & ExpandToken(Var) & " += " & ForStep
  ConvertFor = ConvertFor & ") {"
End Function

Public Function ConvertForEach(ByVal L As String) As String
  Dim Var As String, ForSource As String
  L = Trim(L)
  RecordLeft L, "For "
  RecordLeft L, "Each "
  
  Var = RemoveUntil(L, " In ", True)
  ForSource = L
  
  ConvertForEach = ConvertForEach & "foreach (var iter" & Var & " in " & ConvertExpression(ForSource) & ") {" & vbCrLf & Var & " = iter" & Var & ";"
End Function

Public Function ConvertType(ByVal L As String) As String
  Dim isPrivate As Boolean, isPublic As Boolean
  isPublic = RecordLeft(L, "Public ")
  isPrivate = RecordLeft(L, "Private ")
  RecordLeft L, "Type "
  CurrentTypeName = L
  
  ConvertType = ""
  If Not isPrivate Then ConvertType = ConvertType & "public "
  ConvertType = ConvertType & "class " ' `struct ` is available, but leads to non-conforming behavior when indexing in lists...
  ConvertType = ConvertType & L
  ConvertType = ConvertType & "{ "
End Function

Public Function ConvertTypeLine(ByVal L As String, ByVal vCodeType As CodeType) As String
  ConvertTypeLine = ConvertDeclaration(L, DECL_TYPE, vCodeType)
End Function

Public Function ConvertEnum(ByVal L As String) As String
  Dim isPrivate As Boolean, isPublic As Boolean
  isPublic = RecordLeft(L, "Public ")
  isPrivate = RecordLeft(L, "Private ")
  RecordLeft L, "Enum "
  CurrentEnumName = L
  
  ConvertEnum = ""
  If Not isPrivate Then ConvertEnum = ConvertEnum & "public "
  ConvertEnum = ConvertEnum & "enum "
  ConvertEnum = ConvertEnum & L
  ConvertEnum = ConvertEnum & "{ "
End Function

Public Function ConvertEnumLine(ByVal L As String) As String
  Dim Name As String, Value As String
  Dim Parts() As String
  
  If Trim(L) = "" Then Exit Function
  Parts = Split(L, " = ")
  Name = Trim(Parts(0))
  If UBound(Parts) >= 1 Then Value = Trim(Parts(1)) Else Value = ""
    
  ConvertEnumLine = ""
  If Right(CurrentEnumName, 1) = "+" Then ConvertEnumLine = ConvertEnumLine & ", "
  ConvertEnumLine = ConvertEnumLine & Name
  If Value <> "" Then ConvertEnumLine = ConvertEnumLine & " = " & ConvertExpression(Value)
  CurrentEnumName = CurrentEnumName & "+"  ' convenience
End Function

Public Function ConvertProperty(ByVal L As String, ByVal FullContents As String, ByVal vCodeType As CodeType) As String
  Dim Name As String, IX As Long, isPrivate As Boolean, ReturnType As String, Discard As String
  Dim PropertyType As String
  Dim GetContents As String, SetContents As String
  IX = InStr(L, "(")
  Name = Left(L, IX - 1)
  RecordLeft L, "Public "
  isPrivate = RecordLeft(L, "Private ")
  RecordLeft L, "Property "
  RecordLeft L, "Get "
  RecordLeft L, "Let "
  RecordLeft L, "Set "
  
  IX = InStr(L, "(")
  Name = Left(L, IX - 1)
  If InStr(ModuleProperties, Name) > 0 Then Exit Function
  CurrentFunctionName = Name
  CurrentFunctionReturnValue = "_" & Name
  ModuleProperties = ModuleProperties & "[" & Name & "]"
  
  GetContents = FindPropertyBody(FullContents, "Get", Name, ReturnType)
  If GetContents <> "" Then GetContents = ConvertContents(GetContents, vCodeType, True)
  If ReturnType = "" Then ReturnType = "Variant"
  SetContents = FindPropertyBody(FullContents, "Let", Name, Discard)
  If SetContents = "" Then SetContents = FindPropertyBody(FullContents, "Set", Name, Discard)
  If SetContents <> "" Then SetContents = ConvertContents(SetContents, vCodeType, True)
  
  PropertyType = ConvertArgType(Name, ReturnType)
  
  ConvertProperty = ""
  ConvertProperty = ConvertProperty & IIf(isPrivate, "private ", "public ")
  ConvertProperty = ConvertProperty & IIf(vCodeType = CODE_MODULE, "static ", "")
  ConvertProperty = ConvertProperty & PropertyType & " " & Name & "{ " & vbCrLf
  If GetContents <> "" Then
    ConvertProperty = ConvertProperty & "get {" & vbCrLf
    ConvertProperty = ConvertProperty & PropertyType & " " & CurrentFunctionReturnValue & " = default(" & PropertyType & ");" & vbCrLf
    ConvertProperty = ConvertProperty & GetContents
    ConvertProperty = ConvertProperty & "return " & CurrentFunctionReturnValue & ";" & vbCrLf
    ConvertProperty = ConvertProperty & "}" & vbCrLf
  End If
  If SetContents <> "" Then
    ConvertProperty = ConvertProperty & "set {" & vbCrLf
    ConvertProperty = ConvertProperty & SetContents
    ConvertProperty = ConvertProperty & "}" & vbCrLf
  End If
  ConvertProperty = ConvertProperty & "}" & vbCrLf
End Function

Public Function FindPropertyBody(ByVal FullContents As String, ByVal Typ As String, ByVal Name As String, ByRef ReturnType As String) As String
  Dim X As Long
  X = InStr(FullContents, "Property " & Typ & " " & Name)
  If X = 0 Then Exit Function
  FindPropertyBody = Mid(FullContents, X)
  X = RegExNPos(FindPropertyBody, "\bEnd Property\b", 0)
  FindPropertyBody = Trim(Left(FindPropertyBody, X - 1))
  
  RecordLeft FindPropertyBody, "Property " & Typ & " " & Name
  RecordLeft FindPropertyBody, "("
  X = 1
  Do While X > 0
    If Left(FindPropertyBody, 1) = "(" Then X = X + 1
    If Left(FindPropertyBody, 1) = ")" Then X = X - 1
    FindPropertyBody = Mid(FindPropertyBody, 2)
  Loop
  FindPropertyBody = Trim(FindPropertyBody)
  If StartsWith(FindPropertyBody, "As ") Then
    FindPropertyBody = Mid(FindPropertyBody, 4)
    X = SearchLeft(1, FindPropertyBody, ": " & vbCrLf, False, False)
    ReturnType = Left(FindPropertyBody, X - 1)
    FindPropertyBody = Mid(FindPropertyBody, X)
  End If
  Do While StartsWith(FindPropertyBody, vbCrLf): FindPropertyBody = Mid(FindPropertyBody, 3): Loop
  Do While Right(FindPropertyBody, 2) = vbCrLf: FindPropertyBody = Left(FindPropertyBody, Len(FindPropertyBody) - 2): Loop

  If StartsWith(FindPropertyBody, ":") Then FindPropertyBody = Trim(Mid(FindPropertyBody, 2))
  If Right(FindPropertyBody, 1) = ":" Then FindPropertyBody = Trim(Left(FindPropertyBody, Len(FindPropertyBody) - 1))
End Function

Public Function ConvertDeclaration(ByVal L As String, ByVal declType As DeclarationType, ByVal vCodeType As CodeType) As String
  Dim IsDim As Boolean, isPrivate As Boolean, isPublic As Boolean, IsConst As Boolean, isGlobal As Boolean, isStatic As Boolean
  Dim IsOptional As Boolean, IsByVal As Boolean, IsByRef As Boolean, IsParamArray As Boolean
  Dim IsWithEvents As Boolean, IsEvent As Boolean
  Dim FixedLength As Long, IsNewable As Boolean
  L = Trim(L)
  If L = "" Then Exit Function
  
  IsDim = RecordLeft(L, "Dim ")
  isPrivate = RecordLeft(L, "Private ")
  isPublic = RecordLeft(L, "Public ")
  isGlobal = RecordLeft(L, "Global ")
  IsConst = RecordLeft(L, "Const ")
  isStatic = RecordLeft(L, "Static ")
'  If IsInStr(L, "LineStrings") Then Stop

  If isStatic And declType = DECL_LOCAL Then LineComment = LineComment & " TODO: (NOT SUPPORTED) C# Does not support static local variables."

  Dim Item As Variant, LL As String
  For Each Item In Split(L, ", ")
    Dim IX As Long, ArgName As String, ArgType As String, ArgDefault As String, IsArray As Boolean, IsReferencableType As Boolean
    Dim ArgTargetType As String
    Dim StandardEvent As Boolean
    If ConvertDeclaration <> "" And declType <> DECL_SIGNATURE And declType <> DECL_EXTERN Then ConvertDeclaration = ConvertDeclaration & vbCrLf
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
    
    If StartsWith(ArgType, "New ") Then
      IsNewable = True
      RecordLeft ArgType, "New "
      LineComment = LineComment & "TODO: (NOT SUPPORTED) Dimmable 'New' not supported on variable declaration.  Instantiated only on declaration.  Please ensure usages"
    End If
    
    If InStr(ArgType, " * ") > 0 Then
      FixedLength = Val(Trim(Mid(ArgType, InStr(ArgType, " * ") + 3)))
      ArgType = RemoveUntil(ArgType, " * ")
      LineComment = LineComment & "TODO: (NOT SUPPORTED) Fixed Length String not supported: " & ArgName & "(" & FixedLength & ")"
    End If

    ArgTargetType = ConvertArgType(ArgName, ArgType)
    
    ArgName = LL
    If Right(ArgName, 2) = "()" Then
      IsArray = True
      ArgName = Left(ArgName, Len(ArgName) - 2)
    ElseIf RegExTest(ArgName, "^[a-zA-Z_][a-zA-Z_0-9]*\(.* To .*\)$") Then
      IsArray = True
      LineComment = LineComment & " TODO: (NOT SUPPORTED) Array ranges not supported: " & ArgName
      ArgName = RemoveUntil(ArgName, "(")
    Else
      IsArray = False
    End If
    
    IsReferencableType = ArgTargetType = "Recordset" Or ArgTargetType = "Collection"
    
    ArgTargetType = ConvertArgType(ArgName, ArgType)
    
    StandardEvent = IsStandardEvent(ArgName, ArgType)

    Select Case (declType)
      Case DECL_GLOBAL ' global
        If isPublic Or IsDim Then
          ConvertDeclaration = ConvertDeclaration & "public "
          If vCodeType = CODE_MODULE And Not IsConst Then ConvertDeclaration = ConvertDeclaration & "static "
        Else
          ConvertDeclaration = ConvertDeclaration & "public " & IIf(Not IsConst, "static ", "")
        End If
        If IsConst Then ConvertDeclaration = ConvertDeclaration & "const "
        ConvertDeclaration = ConvertDeclaration & IIf(IsArray, "List<" & ArgTargetType & ">", ArgTargetType) & " "
        ConvertDeclaration = ConvertDeclaration & ArgName
        If ArgDefault <> "" Then
          ConvertDeclaration = ConvertDeclaration & " = " & ConvertExpression(ArgDefault)
        Else
          ConvertDeclaration = ConvertDeclaration & " = " & ArgTypeDefault(ArgTargetType, IsArray, IsNewable) ' VB6 always initializes variables on declaration
        End If
        ConvertDeclaration = ConvertDeclaration & ";"
        If IsArray Then ModuleArrays = ModuleArrays & "[" & ArgName & "]"
      Case DECL_LOCAL ' function contents
        ConvertDeclaration = ConvertDeclaration & IIf(IsArray, "List<" & ArgTargetType & ">", ArgTargetType) & " "
        ConvertDeclaration = ConvertDeclaration & ArgName
        If ArgDefault <> "" Then
          ConvertDeclaration = ConvertDeclaration & " = " & ConvertExpression(ArgDefault)
        Else
          ConvertDeclaration = ConvertDeclaration & " = " & ArgTypeDefault(ArgTargetType, IsArray, IsNewable) ' VB6 always initializes variables on declaration
        End If
        ConvertDeclaration = ConvertDeclaration & ";"
        If IsArray Or IsReferencableType Then CurrentFunctionArrays = CurrentFunctionArrays & "[" & ArgName & "]"
        CurrentFunctionArgs = CurrentFunctionArgs & "[" & ArgName & "]"
      Case DECL_SIGNATURE ' sig def
        If ConvertDeclaration <> "" Then ConvertDeclaration = ConvertDeclaration & ", "
        If IsByRef Or Not IsByVal Then ConvertDeclaration = ConvertDeclaration & "ref "
        ConvertDeclaration = ConvertDeclaration & IIf(IsArray, "List<" & ArgTargetType & ">", ArgTargetType) & " "
        ConvertDeclaration = ConvertDeclaration & ArgName
        If ArgDefault <> "" Then ConvertDeclaration = ConvertDeclaration & " = " & ConvertExpression(ArgDefault) ' default on method sig means optional param
        If IsArray Or IsReferencableType Then CurrentFunctionArrays = CurrentFunctionArrays & "[" & ArgName & "]"
        CurrentFunctionArgs = CurrentFunctionArgs & "[" & ArgName & "]"
      Case DECL_TYPE
        ConvertDeclaration = ConvertDeclaration & "public " & ArgTargetType & " " & ArgName & ";"
      Case DECL_ENUM
      Case DECL_EXTERN
        If ConvertDeclaration <> "" Then ConvertDeclaration = ConvertDeclaration & ", "
        If IsByRef Or Not IsByVal Then ConvertDeclaration = ConvertDeclaration & "ref "
        ConvertDeclaration = ConvertDeclaration & IIf(IsArray, "List<" & ArgTargetType & ">", ArgTargetType) & " "
        ConvertDeclaration = ConvertDeclaration & ArgName
    End Select

'    If IsParamArray Then Stop
    If ArgType = "" And Not IsEvent And Not StandardEvent Then
    End If
    If declType = DECL_SIGNATURE Then
      If IsParamArray Then
      Else
        If Not IsByVal And Not IsByRef And Not StandardEvent Then
        End If
      End If
      If IsOptional And IsByRef Then
      End If
      If IsOptional And ArgDefault = "" Then
      End If
    End If
  Next
End Function

'Function IsStandardEvent(ByVal ArgName As String, ByVal ArgType As String) As Boolean
'  If ArgName = "Cancel" Then IsStandardEvent = True: Exit Function
'  If ArgName = "LastRow" Then IsStandardEvent = True: Exit Function
'  If ArgName = "LastCol" Then IsStandardEvent = True: Exit Function
'  If ArgName = "newCol" Then IsStandardEvent = True: Exit Function
'  If ArgName = "newCol" Then IsStandardEvent = True: Exit Function
'  If ArgName = "newRow" Then IsStandardEvent = True: Exit Function
'  If ArgName = "OldValue" Then IsStandardEvent = True: Exit Function
'  If ArgName = "Index" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
'  If ArgName = "Offset" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
'  If ArgName = "UnloadMode" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
'  If ArgName = "KeyCode" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
'  If ArgName = "KeyAscii" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
'  If ArgName = "Button" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
'  If ArgName = "Shift" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
'  If ArgName = "X" And ArgType = "Single" Then IsStandardEvent = True: Exit Function
'  If ArgName = "Y" And ArgType = "Single" Then IsStandardEvent = True: Exit Function
'  If ArgName = "Source" And ArgType = "Control" Then IsStandardEvent = True: Exit Function
'  If ArgName = "Item" And ArgType = "Integer" Then IsStandardEvent = True: Exit Function
'  IsStandardEvent = False
'End Function
'
Public Function ConvertArgType(ByVal Name As String, ByVal Typ As String) As String
  Select Case Typ
    Case "Long", "Integer", "Int32", "Short":
      ConvertArgType = "int"
    Case "Currency"
      ConvertArgType = "decimal"
    Case "Date"
      ConvertArgType = "DateTime"
    Case "Double", "Float", "Single"
      ConvertArgType = "decimal"
    Case "String":
      ConvertArgType = "string"
    Case "Boolean"
      ConvertArgType = "bool"
    Case "Variant", "Object"
      ConvertArgType = "dynamic"
    Case Else
      ConvertArgType = Typ
  End Select
End Function

Public Function ArgTypeDefault(ByVal ArgType As String, Optional ByVal asArray As Boolean = False, Optional ByVal IsNewable As Boolean = False) As String
  If Not asArray Then
    Select Case LCase(ArgType)
      Case "string"
        ArgTypeDefault = """"""
      Case "long", "int", "integer", "short", "byte", "decimal", "float", "double", "currency"
        ArgTypeDefault = "0"
      Case "boolean", "bool"
        ArgTypeDefault = "false"
      Case "vbtristate"
        ArgTypeDefault = "vbUseDefault"
      Case "datetime", "date"
        ArgTypeDefault = "DateTime.MinValue"
      Case Else
        ArgTypeDefault = IIf(IsNewable, "new " & ArgType & "()", "null")
    End Select
  Else
    ArgTypeDefault = "new List<" & ArgType & ">()"
  End If
End Function

Public Function ConvertSignature(ByVal LL As String, Optional ByVal vCodeType As CodeType = CODE_FORM) As String
  Dim L As String, WithReturn As Boolean
  Dim isPublic As Boolean, isPrivate As Boolean, IsFriend As Boolean
  Dim IsPropertyGet As Boolean, IsPropertyLet As Boolean, IsPropertySet As Boolean
  Dim IsFunction As Boolean, IsSub As Boolean
  L = LL
  isPrivate = RecordLeft(L, "Private ")
  isPublic = RecordLeft(L, "Public ")
  IsFriend = RecordLeft(L, "Friend ")
  IsSub = RecordLeft(L, "Sub ")
  IsFunction = RecordLeft(L, "Function ")
  IsPropertyGet = RecordLeft(L, "Property Get ")
  IsPropertyLet = RecordLeft(L, "Property let ")
  IsPropertySet = RecordLeft(L, "Property set ")
  WithReturn = IsFunction Or IsPropertyGet

  Dim IX As Long, Ix2 As Long, Name As String, Args As String, Ret As String, RetTargetType As String, IsArray As Boolean
  IX = InStr(L, "(")
  If IX = 0 Then Exit Function
  Name = Left(L, IX - 1)
  If RegExTest(L, "\) As .*\(\)$") Then
    Ix2 = InStrRev(L, ")", Len(L) - 2)
  Else
    Ix2 = InStrRev(L, ")")
  End If
  Args = Mid(L, IX + 1, Ix2 - IX - 1)
  Ret = Mid(L, Ix2 + 1)
  Ret = Replace(Ret, " As ", "")
  IsArray = Right(Ret, 2) = "()"
  If IsArray Then Ret = Left(Ret, Len(Ret) - 2)
  RetTargetType = ConvertArgType(Name, Ret)
  If IsArray Then RetTargetType = "List<" & RetTargetType & ">"
  
  CurrentFunctionName = Name
  CurrentFunctionReturnValue = IIf(WithReturn, "_" & CurrentFunctionName, "")
  
  ConvertSignature = ""
  If isPublic Then ConvertSignature = ConvertSignature & "public "
  If isPrivate Then ConvertSignature = ConvertSignature & "private "
  If vCodeType = CODE_MODULE Then ConvertSignature = ConvertSignature & "static "
  ConvertSignature = ConvertSignature & IIf(Ret = "", "void ", RetTargetType & " ")
  ConvertSignature = ConvertSignature & Name & "(" & ConvertDeclaration(Args, DECL_SIGNATURE, vCodeType) & ") {"
  If WithReturn Then
    ConvertSignature = ConvertSignature & vbCrLf & RetTargetType & " " & CurrentFunctionReturnValue & " = " & ArgTypeDefault(RetTargetType) & ";"
  End If
  
  If IsEvent(Name) Then ConvertSignature = EventStub(Name) & ConvertSignature
End Function

Public Function ConvertDeclare(ByVal L As String) As String
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'[DllImport("User32.dll")]
'public static extern int MessageBox(int h, string m, string c, int type);
  Dim isPrivate As Boolean, isPublic As Boolean, IsFunction As Boolean, IsSub As Boolean
  Dim X As Long
  Dim Name As String, cLib As String, cAlias As String
  Dim Args As String, Ret As String
  L = Trim(L)
  
  isPrivate = RecordLeft(L, "Private ")
  isPublic = RecordLeft(L, "Public ")
  L = StripLeft(L, "Declare ")
  IsFunction = RecordLeft(L, "Function ")
  IsSub = RecordLeft(L, "Sub ")
  
  X = InStr(L, " ")
  Name = Left(L, X - 1)
  L = Mid(L, X + 1)
  
  If RecordLeft(L, "Lib ") Then
    X = InStr(L, " ")
    cLib = Left(L, X - 1)
'    If Left(cLib, 1) = """" Then cLib = Mid(cLib, 2, Len(cLib) - 2)
    L = Mid(L, X + 1)
  End If
  
  If RecordLeft(L, "Alias ") Then
    X = InStr(L, " ")
    cAlias = Left(L, X - 1)
'    If Left(cAlias, 1) = """" Then cAlias = Mid(cAlias, 2, Len(cAlias) - 2)
    L = Mid(L, X + 1)
  End If
  
  X = InStrRev(L, ")")
  Ret = Trim(Mid(L, X + 1))
  Ret = Replace(Ret, "As ", "")
  Args = Mid(L, 2, X - 2)
  
  ConvertDeclare = ""
  ConvertDeclare = ConvertDeclare & "[DllImport(" & cLib & ")]" & vbCrLf
  ConvertDeclare = ConvertDeclare & IIf(isPrivate, "private ", "public ") & "static extern "
  ConvertDeclare = ConvertDeclare & IIf(Ret = "", "void", ConvertArgType("return", Ret)) & " "
  ConvertDeclare = ConvertDeclare & Name & "("
  ConvertDeclare = ConvertDeclare & ConvertDeclaration(Args, DECL_EXTERN, True)
  ConvertDeclare = ConvertDeclare & ");"
End Function

Public Function ConvertFileOpen(ByVal L As String) As String
    'Open pathname For mode [ Access access ] [ lock ] As [ # ] filenumber [ Len = reclength ]
  Dim vPath As String, vMode As String, vAccess As String, vLock As Boolean, vNumber As String, vLen As String
  L = Trim(L)
  RecordLeft L, "Open "
  vPath = RemoveUntil(L, " ", True)
  RecordLeft L, "For "
  vMode = RemoveUntil(L, " ", True)
  If RecordLeft(L, "Access ") Then vAccess = RemoveUntil(L, " ", True)
  vLock = RecordLeft(L, "Lock ")
  RecordLeft L, "As #"
  vNumber = L
'  If RecordLeft(L, "Len = ") Then vLen = L
  
  ConvertFileOpen = "FileOpen(" & vNumber & ", " & vPath & ", VBFileMode(""" & vMode & """), VBFileAccess(""" & vMode & """), VBFileShared(""" & vMode & """), VBFileRecLen(""" & vMode & """))"
End Function

Public Function SplitByComma(ByVal L As String) As String()
  Dim Results() As String, ResultCount As Long
  Dim N As Long, I As Long, C As String, Depth As Long, Part As String
  N = Len(L)
  For I = 1 To N
    C = Mid(L, I, 1)
    If C = "(" Then
      Depth = Depth + 1
      Part = Part & C
    ElseIf Depth > 0 And C = ")" Then
      Depth = Depth - 1
      Part = Part & C
    ElseIf Depth = 0 And (C = "," Or C = ")") Then
      ResultCount = ResultCount + 1
      ReDim Preserve Results(1 To ResultCount)
      Results(ResultCount) = Trim(Part)
      Part = ""
    Else
      Part = Part & C
    End If
  Next
  
  ResultCount = ResultCount + 1
  ReDim Preserve Results(1 To ResultCount)
  Results(ResultCount) = Trim(Part)
  
  SplitByComma = Results
End Function

Public Function FindNextOperator(ByVal L As String) As Long
  Dim N As Long
  N = Len(L)
  For FindNextOperator = 1 To N
    If StartsWith(Mid(L, FindNextOperator), " && ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " || ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " ^^ ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " - ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " + ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " * ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " / ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " < ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " > ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " >= ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " <= ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " != ") Then Exit Function
    If StartsWith(Mid(L, FindNextOperator), " == ") Then Exit Function
  Next
  FindNextOperator = 0
End Function

Public Function ConvertIIf(ByVal L As String) As String
  Dim Parts() As String
  Dim Condition As String, TruePart As String, FalsePart As String
  
  Parts = SplitByComma(Mid(Trim(L), 5, Len(L) - 5))
  Condition = Parts(1)
  TruePart = Parts(2)
  FalsePart = Parts(3)
  
  ConvertIIf = "(" & ConvertExpression(Condition) & " ? " & ConvertExpression(TruePart) & " : " & ConvertExpression(FalsePart) & ")"
End Function

Public Function ConvertStatement(ByVal L As String) As String
  Dim NonCodeLine As Boolean
  L = Trim(L)
  
  If StartsWith(L, "Set ") Then L = Mid(L, 5)
  If WithVar <> "" Then L = Replace(" " & L, " .", " " & WithVar & ".")
  
  If StartsWith(L, "Option ") Then
    ' ignore "Option" directives
    NonCodeLine = True
  ElseIf RegExTest(L, "^[ ]*Exit (Function|Sub|Property)$") Then
    ConvertStatement = ConvertStatement & "return"
    If CurrentFunctionReturnValue <> "" Then ConvertStatement = ConvertStatement & " " & CurrentFunctionReturnValue
  ElseIf RegExTest(L, "^[ ]*Exit (Do|Loop|For|While)$") Then
    ConvertStatement = ConvertStatement & "break"
  ElseIf RegExTest(L, "^[ ]*[^ ]+ = ") Then
    Dim IX As Long, AssignmentTarget As String, AssignmentValue As String
    IX = InStr(L, " = ")
    AssignmentTarget = Trim(Left(L, IX - 1))
    If InStr(AssignmentTarget, "(") > 0 Then AssignmentTarget = ConvertExpression(AssignmentTarget)
    If IsControlRef(AssignmentTarget, ModuleName) Then
'      If InStr(AssignmentTarget, "lblPrg") > 0 Then Stop
      AssignmentTarget = modRefScan.FormControlRepl(AssignmentTarget, ModuleName)
    End If
    If AssignmentTarget = CurrentFunctionName Then AssignmentTarget = CurrentFunctionReturnValue
    AssignmentValue = Mid(L, IX + 3)
    ConvertStatement = AssignmentTarget & " = " & ConvertExpression(AssignmentValue)
  ElseIf RegExTest(L, "^[ ]*Unload ") Then
    L = Trim(L)
    RecordLeft L, "Unload "
    ConvertStatement = IIf(L = "Me", "Unload()", L & ".instance.Unload()")
  ElseIf RegExTest(L, "^[ ]*With") Or RegExTest(L, "^[ ]*End With") Then
'    ConvertStatement = "// TODO: (NOT SUPPORTED): " & L
    NonCodeLine = True
  ElseIf RegExTest(L, "^[ ]*(On Error|Resume) ") Then
    ConvertStatement = "// TODO: (NOT SUPPORTED): " & L
    NonCodeLine = True
  ElseIf RegExTest(L, "^[ ]*ReDim ") Then
    ConvertStatement = "// TODO: (NOT SUPPORTED): " & L
    NonCodeLine = True
  ElseIf RegExTest(L, "^[ ]*Err.Clear") Then
    ConvertStatement = "// TODO: (NOT SUPPORTED): " & L
    NonCodeLine = True
  ElseIf RegExTest(L, "^[ ]*(([a-zA-Z_()0-9.]\.)*)?[a-zA-Z_0-9.]+$") Then ' Method call without parens or args (statement, not expression)
    ConvertStatement = ConvertStatement & L & "()"
  ElseIf RegExTest(L, "^[ ]*(Close|Put|Get|Seek|Input|Print|Line Input) [#]") Then
    Dim FileOp As String, FileOpRest As String
    FileOp = RegExNMatch(L, "^[ ]*(Close|Put|Get|Seek|Inpu|Print|Line Input) [#]", 0)
    FileOp = Trim(Replace(FileOp, "#", ""))
    
    FileOpRest = Trim(Mid(L, InStr(L, "#") + 1))
    FileOpRest = Replace(FileOpRest, ", ,", ", _,")
    
    If FileOp = "Put" Then FileOp = "FilePut": FileOpRest = ReorderParams(FileOpRest, Array(0, 2, 1))
    If FileOp = "Get" Then FileOp = "FileGet": FileOpRest = ReorderParams(FileOpRest, Array(0, 2, 1))
    If FileOp = "Close" Then FileOp = "FileClose"
    If FileOp = "Input" Then FileOp = "Input"
    If FileOp = "Print" Then FileOp = "Print"
    If FileOp = "Line Input" Then
      ConvertStatement = ReorderParams(FileOpRest, Array(1)) & " = LineInput(" & ReorderParams(FileOpRest, Array(0)) & ")"
    Else
      Do While (EndsWith(FileOpRest, ", _")): FileOpRest = Left(FileOpRest, Len(FileOpRest) - 3): Loop
      
      ConvertStatement = FileOp & "(" & FileOpRest & ")"
    End If
    LineComment = LineComment & "TODO: (VERIFY) Verify File Access: " & L
  ElseIf RegExTest(L, "^[ ]*Open .* As #") Then
    ConvertStatement = ConvertFileOpen(L)
    LineComment = LineComment & "TODO: (VERIFY) Verify File Access: " & L
  ElseIf RegExTest(L, "^[ ]*Print #") Then
    ConvertStatement = "VBWriteFile(""" & Replace(Trim(L), """", """""") & """)"
    LineComment = LineComment & "TODO: (NOT SUPPORTED) VB File Access Suppressed.  Convert manually: " & L
  ElseIf RegExTest(L, "^[ ]*(([a-zA-Z_()0-9.]\.)*)?[a-zA-Z_0-9.]+ .*") Then ' Method call without parens but with args (statement, not expression)
    Dim FunctionCall As String, ArgList As String, ArgPart As Variant, ArgN As Long
    FunctionCall = RegExNMatch(L, "^[ ]*((([a-zA-Z_()0-9.]\.)*)?[a-zA-Z_0-9.]+)", 0)
    ArgList = Trim(Mid(L, Len(FunctionCall) + 1))
    ConvertStatement = ExpandFunctionCall(FunctionCall, ArgList)
  Else
    ConvertStatement = L
  End If

  If Not NonCodeLine Then ConvertStatement = ConvertStatement & ";"
End Function

Public Function ConvertExpression(ByVal L As String) As String
  L = Replace(L, " \ ", " / ")
  L = Replace(L, " = ", " == ")
  L = Replace(L, " Mod ", " % ")
  L = Replace(L, " & ", " + ")
  L = Replace(L, " And ", " && ")
  L = Replace(L, " Or ", " || ")
  L = Replace(L, " Xor ", " ^^ ")
  L = Replace(L, " Is ", " == ")
  If InStr(L, " Like ") > 0 Then LineComment = LineComment & "TODO: (NOT SUPPORTED) LIKE statement changed to ==: " & L
  L = Replace(L, " Like ", " == ")
  L = Replace(L, " <> ", " != ")
  L = RegExReplace(L, "\bNot\b", "!")
  
  L = RegExReplace(L, "\bFalse\b", "false")
  L = RegExReplace(L, "\bTrue\b", "true")
  
  If LMatch(LTrim(L), "New ") Then L = "new " & Mid(LTrim(L), 5) & "()"
  
  If StartsWith(L, "IIf(") Then
    L = ConvertIIf(L)
  Else
    L = ParseAndExpandExpression(L)
  End If
  
  If CurrentFunctionName <> "" Then L = RegExReplace(L, "\b" & CurrentFunctionName & "([^(a-zA-Z_])", CurrentFunctionReturnValue & "$1")
  
  ConvertExpression = L
End Function

Public Function ParseAndExpandExpression(ByVal Src As String) As String
  Dim S As String, Token As String, T As String
  Dim I As Long, J As Long
  Dim X As Long, Y As Long, C As String
  Dim FunctionName As String, FunctionArgs As String
  
  Token = EXPRESSION_TOKEN_PREFIX & CLng(Rnd * 1000000)
  
  
  S = RegExNMatch(Src, "\([^()]+\)", 0)
  If S <> "" Then
    X = InStr(Src, S)
    Src = Replace(Src, S, Token, 1, 1)
    If X > 1 Then C = Mid(Src, X - 1, 1) Else C = ""
    If X > 1 And C <> "(" And C <> ")" And C <> " " Then
      Y = SearchLeft(X - 1, Src, "() ", False, True)
      FunctionName = Mid(Src, Y + 1, X - Y - 1)
      Src = Replace(Src, FunctionName & Token, Token, 1, 1)
      FunctionArgs = Mid(S, 2, Len(S) - 2)
      If modRefScan.IsControlRef(FunctionName, ModuleName) Then
        ParseAndExpandExpression = FunctionName & "[" & FunctionArgs & "]" & "." & ConvertControlProperty("", "", FormControlRefDeclType(FunctionName, ModuleName))
        Exit Function
      End If
      FunctionName = ExpandToken(FunctionName, True)
      S = ExpandFunctionCall(FunctionName, FunctionArgs)
    
      ParseAndExpandExpression = ParseAndExpandExpression(Src)
      ParseAndExpandExpression = Replace(ParseAndExpandExpression, Token, S)
'      Debug.Print "FUNCTION: " & S
      Exit Function
    Else ' not a function, but sub expression maybe math
      T = Mid(S, 2, Len(S) - 2)
      X = FindNextOperator(T)
      If X = 0 Then
        ParseAndExpandExpression = ExpandToken(T)
      Else
        Y = InStr(X + 2, T, " ")
        S = ExpandToken(Left(T, X - 1)) & Mid(T, X, Y - X + 1) & ParseAndExpandExpression(Mid(T, Y + 1))
      End If
      
      ParseAndExpandExpression = ParseAndExpandExpression(Src)
      ParseAndExpandExpression = Replace(ParseAndExpandExpression, Token, "(" & S & ")")
      
'      Debug.Print "PLAIN EXP: " & S
      Exit Function
    End If
  End If
    
  ' no subexpression.  Check for math
  X = FindNextOperator(Src)
  If X = 0 Then
    ParseAndExpandExpression = ExpandToken(Src)
'    Debug.Print "SIMPLE TOKEN: " & S
    Exit Function
  Else
    Y = InStr(X + 2, Src, " ")
    ParseAndExpandExpression = ParseAndExpandExpression(Left(Src, X - 1)) & Mid(Src, X, Y - X + 1) & ParseAndExpandExpression(Mid(Src, Y + 1))
'    Debug.Print "PARSED ARITHMATIC: " & S
    Exit Function
  End If
End Function

Public Function ExpandToken(ByVal T As String, Optional ByVal WillAddParens As Boolean = False, Optional ByVal AsLast As Boolean = False) As String
  Dim WithNot As Boolean
  WithNot = Left(T, 1) = "!"
  If WithNot Then T = Mid(T, 2)
'  If IsInStr(T, "modConfig") Then Stop
'  If IsInStr(T, "ConfigValid") Then Stop
'  If IsInStr(T, "FreeFile") Then Stop
  
'  Debug.Print "ExpandToken: " & T
  If T = CurrentFunctionName Then
    T = CurrentFunctionReturnValue
  ElseIf T = "Rnd" Then
    T = T & "()"
  ElseIf T = "FreeFile" Then
    T = T & "()"
  ElseIf T = "Me" Then
    T = "this"
  ElseIf T = "App.Path" Then
    T = "AppContext.BaseDirectory"
  ElseIf T = "Now" Then
    T = "DateTime.Now"
  ElseIf T = "Nothing" Then
    T = "null"
  ElseIf T = "Err.Number" Then
    T = "Err().Number"
  ElseIf T = "Err.Description" Then
    T = "Err().Description"
  ElseIf InStr(CurrentFunctionArgs, T) = 0 And Not WillAddParens And (IsFuncRef(T) Or IsLocalFuncRef(T)) Then
'    Debug.Print "Autofunction: " & T
    T = T & "()"
  ElseIf modRefScan.IsFormRef(T) Then
    T = FormRefRepl(T)
  ElseIf modRefScan.IsControlRef(T, ModuleName) Then
    T = FormControlRepl(T, ModuleName)
  ElseIf modRefScan.IsEnumRef(T) Then
    T = modRefScan.EnumRefRepl(T)
  ElseIf Left(T, 2) = "&H" Then  ' hex number
    T = "0x" & Mid(T, 3)
    If Right(T, 1) = "&" Then T = Left(T, Len(T) - 1)
  ElseIf RegExTest(T, "^[0-9.-]+[%&@!#]?$") Then ' plain number.  Maybe:  negative, decimals, or typed
    If RegExTest(T, "^[0-9.-]+[%&@!#]$") Then T = Left(T, Len(T) - 1)
    If IsInStr(T, ".") Then T = T & "m"
  ElseIf IsInStr(T, ".") Then
    Dim Parts() As String, I As Long, Part As String, IsLast As Boolean
    Dim TOut As String
'    Debug.Print "Reference: " & T
    TOut = ""
    Parts = Split(T, ".")
    For I = LBound(Parts) To UBound(Parts)
      Part = Parts(I)
      IsLast = I = UBound(Parts)
      If TOut <> "" Then TOut = TOut & "."
      TOut = TOut & ExpandToken(Part, WillAddParens, IsLast)
    Next
    T = TOut
  End If
  ExpandToken = IIf(WithNot, "!", "") & T
End Function

Public Function ExpandFunctionCall(ByVal FunctionName As String, ByVal Args As String) As String
  If InStr(ModuleArrays & CurrentFunctionArrays & FormControlArrays, "[" & FunctionName & "]") > 0 Then
    ExpandFunctionCall = FunctionName & "[" & ProcessFunctionArgs(Args) & "]"
  ElseIf FunctionName = "LBound" Then
    ExpandFunctionCall = "0"
  ElseIf FunctionName = "UBound" Then
    ExpandFunctionCall = Args & ".Count"
  ElseIf FunctionName = "Split" Then
    ExpandFunctionCall = "new List<string>(" & FunctionName & "(" & ProcessFunctionArgs(Args) & ")" & ")"
  ElseIf FunctionName = "Debug.Print" Then
    ExpandFunctionCall = "Console.WriteLine(" & ProcessFunctionArgs(Args) & ")"
  ElseIf FunctionName = "Erase" Then
    ExpandFunctionCall = Args & ".Clear()"
  ElseIf FunctionName = "GoTo" Then
    ExpandFunctionCall = "goto " & Args
  ElseIf FunctionName = "Array" Then
    ExpandFunctionCall = "new List<dynamic>() {" & ProcessFunctionArgs(Args) & "}"
  ElseIf FunctionName = "Show" Then
    ExpandFunctionCall = IIf(Args = "", "Show()", "ShowDialog()")
  ElseIf modRefScan.IsFormRef(FunctionName) Then
    ExpandFunctionCall = modRefScan.FormRefRepl(FunctionName) & "(" & ProcessFunctionArgs(Args, FunctionName) & ")"
  Else
    ExpandFunctionCall = FunctionName & "(" & ProcessFunctionArgs(Args, FunctionName) & ")"
  End If
  
  ExpandFunctionCall = RegExReplace(ExpandFunctionCall, "\.Show\(.+\)", ".ShowDialog()")
End Function

Public Function ProcessFunctionArgs(ByVal Args As String, Optional ByVal FunctionName As String = "") As String
  Dim Arg As Variant, I As Long
  For Each Arg In SplitByComma(Args)
    I = I + 1
    If ProcessFunctionArgs <> "" Then ProcessFunctionArgs = ProcessFunctionArgs & ", "
    If FunctionName <> "" Then
      If modRefScan.IsFuncRef(FunctionName) Then
        If I <= FuncRefDeclArgCnt(FunctionName) And modRefScan.FuncRefArgByRef(FunctionName, I) Then
'          If IsInStr(Arg, STRING_TOKEN_PREFIX) Then Stop
          ProcessFunctionArgs = ProcessFunctionArgs & "ref "
        End If
      End If
    End If
    cValP Nothing, "", ""
    ProcessFunctionArgs = ProcessFunctionArgs & ConvertExpression(Arg)
  Next
End Function

