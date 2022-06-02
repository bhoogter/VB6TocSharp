using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.DateAndTime;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modProjectFiles;
using static modRegEx;
using static modTextFiles;
using static modUtils;
using static VBExtension;



static class modQuickLint
{
    // Define some constants for easier access.
    public const int MAX_ERRORS_DEFAULT = 50;
    public const string Attr = "Attribute";
    public const string Q = "\"";
    public const string A = "'";
    public const string S = " ";
    public const string LintKey = "'@NO-LINT";
    // Represents all lint types.  If this is disabled, all are disabled.
    public const string TY_ALLTY = "AllTy";
    // Lint Types
    public const string TY_ERROR = "Error";
    public const string TY_INDNT = "Indnt";
    public const string TY_ARGNA = "ArgNa";
    public const string TY_ARGTY = "ArgTy";
    public const string TY_FSPNA = "FSPNa";
    public const string TY_DEPRE = "Depre";
    public const string TY_MIGRA = "Migra";
    public const string TY_STYLE = "Style";
    public const string TY_BLANK = "Blank";
    public const string TY_EXPLI = "Expli";
    public const string TY_COMPA = "Compa";
    public const string TY_TYPEC = "TypeC";
    public const string TY_NOTYP = "NoTyp";
    public const string TY_BYRFV = "ByReV";
    public const string TY_PRIPU = "PriPu";
    public const string TY_FNCRE = "FncRe";
    public const string TY_CORRE = "Corre";
    public const string TY_GOSUB = "GoSub";
    public const string TY_CSTOP = "CStop";
    public const string TY_OPDEF = "OpDef";
    public const string TY_OPBYR = "OpByR";
    public const string TY_DFCTL = "DfCtl";
    // Basic Lint customization here.  Just a comma-separated list of the types above.
    public const int Idnt = 2; // Set to your preferred indent.  Default is 4.  We always used 2.
    public const string DISABLED_LINT_TYPES = TY_OPBYR; // TY_ARGTY & __S1 & TY_OPDEF
    public const string WARNING_LINT_TYPES = "";
    public const string AUTOFIX_LINT_TYPES = TY_INDNT + "," + TY_ARGNA + "," + TY_OPDEF + "," + TY_NOTYP + "," + TY_STYLE;
    public static string ErrorPrefix = ""; // Just a module global to not have to calculate this each time.  Prepends each lint error.
    public static string ErrorIgnore = ""; // Any error types in this string are ignored.
    public static List<string> AutofixFind = new List<string>();
    public static List<string> AutofixRepl = new List<string>();
    public static List<string> AutofixFindRestOfFile = new List<string>();
    public static List<string> AutofixReplRestOfFile = new List<string>();
    public static Collection WellKnownNames = new Collection(); // Used for lint fixing.  Define well known name to always have `cmd` lint to `Cmd`, or whatever other definitions.TODO: (NOT SUPPORTED) Dimmable 'New' not supported on variable declaration.  Instantiated only on declaration.  Please ensure usages
    public static List<String> ErrorTypes()
    {
        return new List<String>() { TY_ALLTY, TY_ERROR, TY_INDNT, TY_ARGNA, TY_ARGTY, TY_FSPNA, TY_DEPRE, TY_MIGRA, TY_STYLE, TY_BLANK, TY_EXPLI, TY_COMPA, TY_TYPEC, TY_NOTYP, TY_BYRFV, TY_PRIPU, TY_FNCRE, TY_CORRE, TY_GOSUB, TY_CSTOP, TY_OPDEF, TY_OPBYR, TY_DFCTL };
    }
    private static string ResolveSources(string FileName)
    {
        string _ResolveSources = "";
        if (FileName == "") FileName = "prj.vbp";
        if (FileName == "forms")
        {
            _ResolveSources = VBPForms("");
        }
        else if (FileName == "modules")
        {
            _ResolveSources = VBPModules();
        }
        else if (FileName == "classes")
        {
            _ResolveSources = VBPClasses();
        }
        else if (FileName == "usercontrols")
        {
            _ResolveSources = VBPUserControls();
        }
        else
        {
            if (InStr(FileName, "\\") == 0) FileName = AppContext.BaseDirectory + "\\" + FileName;
            _ResolveSources = (Right(FileName, 4) == ".vbp" ? VBPCode(FileName) : FileName);
        }
        return _ResolveSources;
    }
    public static string Lint(string FileName = "", bool Alert = false)
    {
        string _Lint = "";
        string FileList = "";
        FileList = ResolveSources(FileName);
        _Lint = QuickLintFiles(FileList, MAX_ERRORS_DEFAULT);
        if (Alert && _Lint != "") MsgBox(_Lint);
        return _Lint;
    }
    public static string LintFix(string FileName = "")
    {
        string _LintFix = "";
        string FileList = "";
        FileList = ResolveSources(FileName);
        QuickLintFiles(FileList, 0, true);
        return _LintFix;
    }
    public static string QuickLintFiles(string List, int MaxErrors = MAX_ERRORS_DEFAULT, bool AutoFix = false)
    {
        string _QuickLintFiles = "";
        int lintDotsPerRow = 50;
        dynamic L = null;
        int X = 0;
        DateTime StartTime = DateTime.MinValue;
        StartTime = DateTime.Now;
        foreach (var iterL in new List<string>(Split(List, vbCrLf)))
        {
            L = iterL;
            string Result = "";
            if (Trim(L) == "") goto NextFile;
            Result = QuickLintFile(L, MaxErrors, AutoFix);
            if (Result != "")
            {
                string S = "";
                Console.WriteLine(vbCrLf + "Done (" + DateDiff("s", StartTime, DateTime.Now) + "s).   To re-run for failing file, hit enter on the line below:");
                S = "LINT FAILED: " + L + vbCrLf + Result + vbCrLf + "?Lint(\"" + L + "\")";
                _QuickLintFiles = S;
                return _QuickLintFiles;
            }
            else
            {
                Console.WriteLine(Switch(Right(L, 3) == "frm", "o", Right(L, 3) == "cls", "x", true, "."));
            }
            X = X + 1;
            if (X >= lintDotsPerRow) { X = 0; Console.WriteLine(); }
        NextFile:;
            DoEvents();
        }
        Console.WriteLine(vbCrLf + "Done (" + DateDiff("s", StartTime, DateTime.Now) + "s).");
        _QuickLintFiles = "";
        return _QuickLintFiles;
    }
    public static string QuickLintFile(string File, int MaxErrors = MAX_ERRORS_DEFAULT, bool AutoFix = false)
    {
        string _QuickLintFile = "";
        if (InStr(File, "\\") == 0) File = AppContext.BaseDirectory + "\\" + File;
        string fName = "";
        string Contents = "";
        string GivenName = "";
        string CheckName = "";
        fName = Mid(File, InStrRev(File, "\\") + 1);
        CheckName = fName;
        CheckName = Replace(CheckName, ".bas", "");
        CheckName = Replace(CheckName, ".frm", "");
        CheckName = Replace(CheckName, ".cls", "");
        CheckName = Replace(CheckName, ".ctl", "");
        ErrorPrefix = Right(Space(18) + fName, 18) + " ";
        Contents = ReadEntireFile(File);
        GivenName = RegExNMatch(Contents, "Attribute VB_Name = \"([^\"]+)\"", 0);
        GivenName = Replace(Replace(GivenName, "Attribute VB_Name = ", ""), "\"", "");
        if (LCase(CheckName) != LCase(GivenName))
        {
            _QuickLintFile = "Module name [" + GivenName + "] must match file name [" + fName + "].  Rename module or class to match the other";
            return _QuickLintFile;
        }
        _QuickLintFile = QuickLintContents(Contents, MaxErrors, IIf(AutoFix, File, ""));
        return _QuickLintFile;
    }
    public static string QuickLintContents(string Contents, int MaxErrors = MAX_ERRORS_DEFAULT, string AutoFix = "")
    {
        string _QuickLintContents = "";
        List<string> Lines = new List<string>();
        dynamic ActualLine = null;
        string LL = "";
        string L = "";
        // TODO: (NOT SUPPORTED): On Error GoTo LintError
        DisableLintType(DISABLED_LINT_TYPES, true);
        Lines = new List<string>(Split(Replace(Contents, vbCr, ""), vbLf));
        AutofixFind.Clear();
        AutofixRepl.Clear();
        AutofixFindRestOfFile.Clear();
        AutofixReplRestOfFile.Clear();
        bool InAttributes = false;
        bool InBody = false;
        string MultiLineOrig = "";
        string MultiLine = "";
        bool IsMultiLine = false;
        int Indent = 0;
        int LineN = 0;
        string Errors = "";
        int ErrorCount = 0;
        int BlankLineCount = 0;
        Collection Options = new Collection(); // TODO: (NOT SUPPORTED) Dimmable 'New' not supported on variable declaration.  Instantiated only on declaration.  Please ensure usages
        string NewContents = "";
        Indent = 0;
        TestDefaultControlNames(ref Errors, ref ErrorCount, 0, Contents);
        foreach (var iterActualLine in Lines)
        {
            ActualLine = iterActualLine;
            LL = ActualLine;
            if (MaxErrors > 0 && ErrorCount >= MaxErrors) break;
            IsMultiLine = false;
            if (Right(LL, 2) == " _")
            {
                string Portion = "";
                Portion = Left(LL, Len(LL) - 2);
                MultiLineOrig = MultiLineOrig + LL + vbCrLf;
                if (MultiLine != "") Portion = " " + Trim(Portion);
                MultiLine = MultiLine + Portion;
                LineN = LineN + 1;
                goto NextLineWithoutRecord;
            }
            else if (MultiLine != "")
            {
                MultiLineOrig = MultiLineOrig + LL;
                LL = MultiLine + " " + Trim(LL);
                MultiLine = "";
                IsMultiLine = true;
            }
            else
            {
                MultiLineOrig = "";
            }
            TestBlankLines(ref Errors, ref ErrorCount, LineN, LL, ref BlankLineCount);
            TestLintControl(LL);
            L = CleanLine(LL);
            if (!InBody)
            {
                bool IsAttribute = false;
                IsAttribute = Left(L, 10) == "Attribute ";
                if (!InAttributes && IsAttribute)
                {
                    InAttributes = true;
                    goto NextLine;
                }
                else if (InAttributes && !IsAttribute)
                {
                    InAttributes = false;
                    InBody = true;
                    LineN = 0;
                }
                else
                {
                    goto NextLine;
                }
            }
            LineN = LineN + 1;
            // If LineN = 487 Then Stop
            bool UnindentedAlready = false;
            if (RegExTest(L, " ^Option ")) Options.Add("true", Replace(L, "Options ", ""));
            if (RegExTest(L, "^[ ]*(Else|ElseIf .* Then)$"))
            {
                Indent = Indent - Idnt;
            }
            else if (RegExTest(L, "^[ ]*End Select$"))
            {
                Indent = Indent - Idnt - Idnt;
            }
            else if (RegExTest(L, "^[ ]*(End (If|Function|Sub|Property|Enum|Type)|Next( .*)?|Wend|Loop|Loop (While .*|Until .*)|ElseIf .*)$"))
            {
                Indent = Indent - Idnt;
                UnindentedAlready = true;
            }
            else
            {
                UnindentedAlready = false;
            }
            int LineIndent = 0;
            LineIndent = 0;
            while (Mid(RTrim(L), LineIndent + 1, 1) == S) { LineIndent = LineIndent + 1; }
            // If LineN = 210 Then Stop
            TestIndent(ref Errors, ref ErrorCount, LineN, L, LineIndent, (!RegExTest(L, "^[ ]*Case ") ? Indent : Indent - Idnt));
            List<string> Statements = new List<string>();
            dynamic SS = null;
            string St = "";
            Statements = new List<string>(Split(L, ": "));
            foreach (var iterSS in Statements)
            {
                SS = iterSS;
                St = SS;
                if (RegExTest(L, "^[ ]*(Else|ElseIf .*)$"))
                {
                    Indent = Indent + Idnt;
                }
                else if (RegExTest(St, "^[ ]*(End (If|Function|Sub|Property)|Loop|Loop .*|Enum|Type|Select)$"))
                {
                    if (!UnindentedAlready) Indent = Indent - Idnt;
                }
                else if (RegExTest(St, "^[ ]*If "))
                {
                    if (!RegExTest(St, "\\bThen "))
                    {
                        Indent = Indent + Idnt;
                    }
                    else
                    {
                        string IfStatementBody = "";
                        IfStatementBody = Mid(L, InStr(L, " Then "));
                        if (RegExTest(IfStatementBody, "\\b(While |For )\\b"))
                        {
                            RecordError(ref Errors, ref ErrorCount, TY_STYLE, LineN, "Place For/While on separate line from If.  Indent check disabled.", TY_INDNT);
                        }
                        else
                        {
                            TestCodeLine(ref Errors, ref ErrorCount, LineN, IfStatementBody);
                        }
                    }
                }
                else if (RegExTest(St, "^[ ]*For "))
                {
                    Indent = Indent + Idnt;
                }
                else if (RegExTest(St, "^[ ]*Next$"))
                {
                    if (!UnindentedAlready) Indent = Indent - Idnt;
                }
                else if (RegExTest(St, "^[ ]*Next [a-zA-Z_][a-zA-Z0-9_]*$"))
                {
                    RecordError(ref Errors, ref ErrorCount, TY_STYLE, LineN, "Remove variable from NEXT statement");
                    AddFix(TY_STYLE, "Next [a-zA-Z_][a-zA-Z0-9_]*$", "Next");
                    if (!UnindentedAlready) Indent = Indent - Idnt;
                }
                else if (RegExTest(St, "^[ ]*While "))
                {
                    RecordError(ref Errors, ref ErrorCount, TY_STYLE, LineN, "Use Do While/Until...Loop in place of While...Wend");
                    AddFix(TY_STYLE, "\\bWhile\\b", "Do While");
                    Indent = Indent + Idnt;
                }
                else if (RegExTest(St, "^[ ]*Wend"))
                {
                    AddFix(TY_STYLE, "\\bWend\\b", "Loop");
                }
                else if (RegExTest(St, "^[ ]*Do (While|Until)"))
                {
                    Indent = Indent + Idnt;
                }
                else if (RegExTest(St, "^[ ]*Loop$"))
                {
                }
                else if (RegExTest(St, "^[ ]*Do$"))
                {
                    Indent = Indent + Idnt;
                }
                else if (RegExTest(St, "^[ ]*Loop While"))
                {
                    Indent = Indent - Idnt;
                }
                else if (RegExTest(St, "^[ ]*Select Case "))
                {
                    Indent = Indent + Idnt + Idnt;
                }
                else if (RegExTest(St, "^[ ]*With "))
                {
                    RecordError(ref Errors, ref ErrorCount, TY_MIGRA, LineN, "Remove all uses of WITH.  No migration path exists.  Indent check disabled.", TY_INDNT);
                }
                else if (RegExTest(St, "^[ ]*(Private |Public )?Declare (Function |Sub )"))
                {
                    // External Api
                }
                else if (RegExTest(St, "^((Private|Public|Friend) )?Function "))
                {
                    if (!RegExTest(St, ": End Function")) Indent = Indent + Idnt;
                    TestSignature(ref Errors, ref ErrorCount, LineN, St);
                }
                else if (RegExTest(St, "^((Private|Public|Friend) )?Sub "))
                {
                    if (!RegExTest(St, ": End Sub")) Indent = Indent + Idnt;
                    TestSignature(ref Errors, ref ErrorCount, LineN, St);
                }
                else if (RegExTest(St, "^((Private|Public|Friend) )?Property (Get|Let|Set) "))
                {
                    if (!RegExTest(St, ": End Property")) Indent = Indent + Idnt;
                    TestSignature(ref Errors, ref ErrorCount, LineN, St);
                }
                else if (RegExTest(St, "^[ ]*(Public |Private )?(Enum |Type )"))
                {
                    Indent = Indent + Idnt;
                }
                else if (RegExTest(St, "^[ ]*(Public |Private )?Declare "))
                {
                    Indent = Indent + Idnt;
                }
                else if (RegExTest(St, "^[ ]*(Dim|Private|Public|Const|Global) "))
                {
                    TestDeclaration(ref Errors, ref ErrorCount, LineN, St, false);
                }
                else
                {
                    TestCodeLine(ref Errors, ref ErrorCount, LineN, St);
                }
            NextStatement:;
            }
        NextLine:;
            if (AutoFix != "")
            {
                string Fixed = "";
                // If IsMultiLine Then Stop
                // If InStr(LL, __S1) > 0 Then Stop
                // If InStr(LL, __S1) > 0 Then Stop
                if (IsMultiLine)
                {
                    Fixed = PerformAutofix(MultiLineOrig);
                }
                else
                {
                    Fixed = PerformAutofix(LL);
                }
                NewContents = NewContents + Fixed + vbCrLf;
            }
        NextLineWithoutRecord:;
        }
        if (AutoFix != "") WriteFile(AutoFix, Left(NewContents, Len(NewContents) - 2), true);
        TestModuleOptions(ref Errors, ref ErrorCount, Options);
        _QuickLintContents = Errors;
        return _QuickLintContents;
    LintError:;
        RecordError(ref Errors, ref ErrorCount, TY_ERROR, 0, "Linter Error [" + Err().Number + "]: " + Err().Description + ".  Actual Line [" + LineN + "]: " + ActualLine);
        _QuickLintContents = Errors;
        // TODO: (NOT SUPPORTED): Resume Next
        return _QuickLintContents;
    }
    private static string ReadEntireFile(string tFileName)
    {
        string _ReadEntireFile = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        dynamic mFSO = null;
        mFSO = CreateObject("Scripting.FileSystemObject");
        _ReadEntireFile = mFSO.OpenTextFile(tFileName, 1).ReadAll;
        if (FileLen(tFileName) / 10 != Len(_ReadEntireFile) / 10)
        {
            MsgBox("ReadEntireFile was short: " + FileLen(tFileName) + " vs " + Len(_ReadEntireFile));
        }
        return _ReadEntireFile;
    }
    public static string CleanLine(string Line)
    {
        string _CleanLine = "";
        int X = 0;
        int Y = 0;
        while (true)
        {
            X = InStr(Line, Q);
            if (X == 0) break;
            Y = InStr(X + 1, Line, Q);
            while (Mid(Line, Y + 1, 1) == Q)
            {
                Y = InStr(Y + 2, Line, Q);
            }
            if (Y == 0) break;
            Line = Left(Line, X - 1) + new String('S', Y - X + 1) + Mid(Line, Y + 1);
        }
        X = InStr(Line, A);
        if (X > 0) Line = RTrim(Left(Line, X - 1));
        _CleanLine = Line;
        return _CleanLine;
    }
    public static void RecordError(ref string Errors, ref int ErrorCount, string Typ, int LineN, string Error, string DisableLintOnError = "")
    {
        string eLine = "";
        if (IsLintTypeDisabled(Typ)) return;
        if (Len(Errors) != 0) Errors = Errors + vbCrLf;
        if (InStr(Join(ErrorTypes().ToArray(), ","), Typ) == 0)
        {
            Errors = Errors + ErrorPrefix + "[" + TY_ERROR + "] Line " + Right(Space(5) + LineN, 5) + ": Unknown error type in linter (add to ErrorTypes): " + Typ;
        }
        eLine = ErrorPrefix + "[" + Right(Space(5) + Typ, 5) + "] Line " + Right(Space(5) + LineN, 5) + ": " + Error;
        if (InStr(WARNING_LINT_TYPES, Typ) > 0)
        {
            Console.WriteLine("WARNING: " + eLine);
        }
        else
        {
            Errors = Errors + eLine;
            ErrorCount = ErrorCount + 1;
        }
        if (DisableLintOnError != "") DisableLintType(DisableLintOnError);
    }
    public static bool StartsWith(string L, string Find)
    {
        bool _StartsWith = false;
        _StartsWith = Left(L, Len(Find)) == Find;
        return _StartsWith;
    }
    public static string StripLeft(string L, string Find)
    {
        string _StripLeft = "";
        if (StartsWith(L, Find)) _StripLeft = Mid(L, Len(Find) + 1); else _StripLeft = L;
        return _StripLeft;
    }
    public static bool RecordLeft(ref string L, string Find)
    {
        bool _RecordLeft = false;
        _RecordLeft = StartsWith(L, Find);
        if (_RecordLeft) L = Mid(L, Len(Find) + 1);
        return _RecordLeft;
    }
    public static void TestIndent(ref string Errors, ref int ErrorCount, int LineN, string L, int LineIndent, int ExpectedIndent)
    {
        if (RTrim(L) == "") return;
        if (RegExTest(L, "^On Error ")) return;
        if (RegExTest(L, "^[a-zA-Z][a-zA-Z0-9]*:$")) return;
        if (RegExTest(L, "#(If|End If|Else|Const)")) return;
        if (StartsWith(L, "Debug.")) return;
        if (!IsLintTypeDisabled(TY_INDNT))
        {
            if (LineIndent != ExpectedIndent)
            {
                RecordError(ref Errors, ref ErrorCount, TY_INDNT, LineN, "Incorrect Indent -- expected " + ExpectedIndent + ", got " + LineIndent);
                AddFix(TY_INDNT, "^[ ]*", Space(ExpectedIndent));
            }
        }
    }
    public static void TestBlankLines(ref string Errors, ref int ErrorCount, int LineN, string L, ref int BlankLineCount)
    {
        if (Trim(L) != "")
        {
            BlankLineCount = 0;
            return;
        }
        BlankLineCount = BlankLineCount + 1;
        if (BlankLineCount > 3) RecordError(ref Errors, ref ErrorCount, TY_BLANK, LineN, "Too many blank lines.");
    }
    public static void TestLintControl(string L)
    {
        dynamic LL = null;
        if (InStr(L, LintKey) == 0) return;
        string Match = "";
        string Typ = "";
        Match = RegExNMatch(L, LintKey + "(-.....)?$", 0);
        if (Match == "") return;
        Typ = (Match == LintKey ? TY_ALLTY : Replace(Match, LintKey + "-", ""));
        DisableLintType(Typ);
    }
    public static bool IsLintTypeDisabled(string Typ)
    {
        bool _IsLintTypeDisabled = false;
        _IsLintTypeDisabled = InStr(UCase(ErrorIgnore), UCase(Typ)) > 0 || InStr(ErrorIgnore, TY_ALLTY) > 0;
        return _IsLintTypeDisabled;
    }
    public static void DisableLintType(string Typ, bool Reset = false)
    {
        if (Reset) ErrorIgnore = "";
        if (Typ == "") return;
        if (IsNotInStr(Typ, ","))
        {
            if (IsLintTypeDisabled(Typ)) return;
            ErrorIgnore = ErrorIgnore + "," + Typ;
        }
        else
        {
            dynamic L = null;
            foreach (var iterL in new List<string>(Split(Typ, ",")))
            {
                L = iterL;
                DisableLintType(L);
            }
        }
    }
    public static void TestModuleOptions(ref string Errors, ref int ErrorCount, Collection Options)
    {
        // TODO: (NOT SUPPORTED): On Error Resume Next
        string Value = "";
        Value = "";
        Value = (string)Options["Explicit"];
        if (Value != "") RecordError(ref Errors, ref ErrorCount, TY_EXPLI, 0, "Option Explicit not set on file");
        Value = "";
        Value = (string)Options["Compare Binary"];
        Value = (string)Options["Compare Database"];
        if (Value != "") RecordError(ref Errors, ref ErrorCount, TY_COMPA, 0, "Use of Option Compare not recommended");
    }
    public static void TestArgName(ref string Errors, ref int ErrorCount, int LineN, string Name)
    {
        string LL = "";
        LL = Trim(Name);
        if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*%$"))
        { // % Integer Dim L%
            RecordError(ref Errors, ref ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Integer deprecated: " + LL);
            LL = Left(LL, Len(LL) - 1);
            AddFix(TY_TYPEC, "\\b" + LL + ".\\b", LL + " As Integer");
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*&$"))
        { // & Long  Dim M&
            RecordError(ref Errors, ref ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Long deprecated: " + LL);
            LL = Left(LL, Len(LL) - 1);
            AddFix(TY_TYPEC, "\\b" + LL + ".\\b", LL + " As Long");
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*@$"))
        { // @ Decimal Const W@ = 37.5
            RecordError(ref Errors, ref ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Decimal deprecated: " + LL);
            LL = Left(LL, Len(LL) - 1);
            AddFix(TY_TYPEC, "\\b" + LL + ".\\b", LL + " As Decimal");
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-TY_TYPEC-Z0-9_]*!$"))
        { // ! Single  Dim Q!
            RecordError(ref Errors, ref ErrorCount, TY_DEPRE, LineN, "Use of Type Character For Single deprecated: " + LL);
            LL = Left(LL, Len(LL) - 1);
            AddFix(TY_TYPEC, "\\b" + LL + ".\\b", LL + " As Single");
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*#$"))
        { // # Double  Dim X#
            RecordError(ref Errors, ref ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Double deprecated: " + LL);
            LL = Left(LL, Len(LL) - 1);
            AddFix(TY_TYPEC, "\\b" + LL + ".\\b", LL + " As Double");
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*\\$$"))
        { // $ String  Dim V$ = __S2
            RecordError(ref Errors, ref ErrorCount, TY_TYPEC, LineN, "Use of Type Character For String deprecated: " + LL);
            LL = Left(LL, Len(LL) - 1);
            AddFix(TY_TYPEC, "\\b" + LL + ".\\b", LL + " As String");
        }
        if (RegExTest(LL, "^[a-z][a-z0-9_]*[%&@!#$]?$"))
        {
            RecordError(ref Errors, ref ErrorCount, TY_ARGNA, LineN, "Identifier name declared as all lower-case: " + LL);
            AddFix(TY_ARGNA, "\\b" + LL + "\\b", WellKnownName(LL), true);
        }
    }
    public static string WellKnownName(string Str)
    {
        string _WellKnownName = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        InitWellKnownNames();
        _WellKnownName = "";
        _WellKnownName = (string)WellKnownNames[LCase(Str)];
        if (_WellKnownName == "") _WellKnownName = Capitalize(Str);
        return _WellKnownName;
    }
    private static void AddWellKnownName(string S)
    {
        // TODO: (NOT SUPPORTED): On Error Resume Next
        WellKnownNames.Add(S, LCase(S));
    }
    public static void InitWellKnownNames()
    {
        dynamic L = null;
        if (WellKnownNames.Count > 0) return;
        foreach (var iterL in new List<dynamic>() { "hWnd" })
        {
            L = iterL;
            AddWellKnownName(L);
        }
    }
    public static void TestSignatureName(ref string Errors, ref int ErrorCount, int LineN, string Name)
    {
        string LL = "";
        LL = Trim(Name);
        if (RegExTest(LL, "^[a-z][a-z0-9_]*$")) RecordError(ref Errors, ref ErrorCount, TY_FSPNA, LineN, "Func/Sub/Prop name declared as all lower-case: " + LL);
    }
    public static void TestDeclaration(ref string Errors, ref int ErrorCount, int LineN, string L, bool InSignature)
    {
        bool IsOptional = false;
        bool IsByVal = false;
        bool IsByRef = false;
        bool IsParamArray = false;
        bool IsWithEvents = false;
        bool IsEvent = false;
        L = Trim(L);
        L = StripLeft(L, "Dim ");
        L = StripLeft(L, "Private ");
        L = StripLeft(L, "Public ");
        L = StripLeft(L, "Const ");
        L = StripLeft(L, "Global ");
        dynamic Item = null;
        string LL = "";
        foreach (var iterItem in new List<string>(Split(L, ", ")))
        {
            Item = iterItem;
            int IX = 0;
            string ArgName = "";
            string ArgType = "";
            string ArgDefault = "";
            bool StandardEvent = false;
            LL = Item;
            IsEvent = RecordLeft(ref LL, "Event ");
            IsWithEvents = RecordLeft(ref LL, "WithEvents ");
            IsOptional = RecordLeft(ref LL, "Optional ");
            IsByVal = RecordLeft(ref LL, "ByVal ");
            IsByRef = RecordLeft(ref LL, "ByRef ");
            IsParamArray = RecordLeft(ref LL, "ParamArray ");
            IX = InStr(LL, " = ");
            if (IX > 0)
            {
                ArgDefault = Trim(Mid(LL, IX + 3));
                LL = Left(LL, IX - 1);
            }
            else
            {
                ArgDefault = "";
            }
            IX = InStr(LL, " As ");
            if (IX > 0)
            {
                ArgType = Trim(Mid(LL, IX + 4));
                LL = Left(LL, IX - 1);
            }
            else
            {
                ArgType = "";
            }
            ArgName = LL;
            StandardEvent = IsStandardEvent(ArgName, ArgType);
            // If IsParamArray Then Stop
            if (ArgType == "" && !IsEvent && !StandardEvent)
            {
                RecordError(ref Errors, ref ErrorCount, TY_NOTYP, LineN, "Local Parameter Missing Type: [" + ArgName + "]");
                AddFix(TY_NOTYP, "\\b" + ArgName + "\\b", ArgName + " As Variant");
            }
            if (InSignature)
            {
                if (IsParamArray)
                {
                    if (Right(LL, 2) != "()") RecordError(ref Errors, ref ErrorCount, TY_STYLE, LineN, "ParamArray variable not declared as an Array.  Add '()': " + ArgName);
                }
                else
                {
                    if (!IsByVal && !IsByRef && !StandardEvent)
                    {
                        RecordError(ref Errors, ref ErrorCount, TY_BYRFV, LineN, "ByVal or ByRef not specified on parameter [" + ArgName + "] -- specify one or the other");
                        AddFix(TY_BYRFV, "\\b" + Item + "\\b", "ByRef " + Item);
                    }
                }
                if (IsOptional && IsByRef)
                {
                    RecordError(ref Errors, ref ErrorCount, TY_OPBYR, LineN, "Modifiers 'Optional ByRef' may not migrate well: " + ArgName);
                }
                if (IsOptional && ArgDefault == "")
                {
                    RecordError(ref Errors, ref ErrorCount, TY_OPDEF, LineN, "Parameter declared OPTIONAL but no default specified. Must specify default: " + ArgName);
                    AddFix(TY_OPDEF, "\\b" + Item + "\\b", Item + " = " + GetTypeDefault(ArgType));
                }
            }
            TestArgName(ref Errors, ref ErrorCount, LineN, LL);
            if (!StandardEvent) TestArgType(ref Errors, ref ErrorCount, LineN, LL, ArgType);
        }
    }
    public static string GetTypeDefault(string ArgType)
    {
        string _GetTypeDefault = "";
        switch (LCase(ArgType))
        {
            case "string":
                _GetTypeDefault = "\"\"";
                break;
            case "long":
            case "integer":
            case "short":
            case "byte":
            case "date":
            case "decimal":
            case "float":
            case "double":
            case "currency":
                _GetTypeDefault = "0";
                break;
            case "boolean":
                _GetTypeDefault = "False";
                break;
            case "vbtristate":
                _GetTypeDefault = "vbUseDefault";
                break;
            default:
                _GetTypeDefault = "Nothing";
                break;
        }
        return _GetTypeDefault;
    }
    public static bool IsStandardEvent(string ArgName, string ArgType)
    {
        bool _IsStandardEvent = false;
        if (ArgName == "Cancel") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "LastRow") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "LastCol") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "newCol") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "newCol") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "newRow") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "OldValue") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "Index" && ArgType == "Integer") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "Offset" && ArgType == "Integer") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "UnloadMode" && ArgType == "Integer") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "KeyCode" && ArgType == "Integer") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "KeyAscii" && ArgType == "Integer") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "Button" && ArgType == "Integer") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "Shift" && ArgType == "Integer") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "X" && ArgType == "Single") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "Y" && ArgType == "Single") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "Source" && ArgType == "Control") { _IsStandardEvent = true; return _IsStandardEvent; }
        if (ArgName == "Item" && ArgType == "Integer") { _IsStandardEvent = true; return _IsStandardEvent; }
        _IsStandardEvent = false;
        return _IsStandardEvent;
    }
    public static void TestArgType(ref string Errors, ref int ErrorCount, int LineN, string Name, string Typ)
    {
        string Expect = "";
        if (Typ == "Integer") Expect = "Long";
        if (Typ == "Short") Expect = "Long";
        if (Typ == "Byte") Expect = "Long";
        if (Typ == "Float") Expect = "Double";
        if (Typ == "Any") Expect = "String";
        if (Expect != "")
        {
            RecordError(ref Errors, ref ErrorCount, TY_ARGTY, LineN, "Arg [" + Name + "] is of type [" + Typ + "] -- use " + Expect + " (or disable type linting for file)");
        }
    }
    public static void TestSignature(ref string Errors, ref int ErrorCount, int LineN, string LL)
    {
        if (!RegExTest(LL, "^[ ]*(Private|Public|Friend) ")) RecordError(ref Errors, ref ErrorCount, TY_PRIPU, LineN, "Either Private or Public should be specified, but neither was.");
        string L = "";
        bool WithReturn = false;
        L = LL;
        L = StripLeft(L, "Private ");
        L = StripLeft(L, "Public ");
        L = StripLeft(L, "Friend ");
        L = StripLeft(L, "Sub ");
        if (StartsWith(L, "Function ") || StartsWith(L, "Property Get ")) WithReturn = true;
        L = StripLeft(L, "Function ");
        L = StripLeft(L, "Property Get ");
        L = StripLeft(L, "Property Let ");
        L = StripLeft(L, "Property Set ");
        int IX = 0;
        int Ix2 = 0;
        string Name = "";
        string Args = "";
        string Ret = "";
        IX = InStr(L, "(");
        if (IX == 0) return;
        Name = Left(L, IX - 1);
        if (RegExTest(L, "\\) As .*\\(\\)$"))
        {
            Ix2 = InStrRev(L, ")", Len(L) - 2);
        }
        else
        {
            Ix2 = InStrRev(L, ")");
        }
        Args = Mid(L, IX + 1, Ix2 - IX - 1);
        Ret = Mid(L, Ix2 + 1);
        TestSignatureName(ref Errors, ref ErrorCount, LineN, Name);
        if (WithReturn && Ret == "") RecordError(ref Errors, ref ErrorCount, TY_FNCRE, LineN, "Function Return Type Not Specified -- Specify Return Type or Variant");
        TestDeclaration(ref Errors, ref ErrorCount, LineN, Args, true);
    }
    public static void TestDefaultControlNames(ref string Errors, ref int ErrorCount, int LineN, string Contents)
    {
        List<dynamic> vTypes = new List<dynamic>();
        dynamic vType = null;
        string Matcher = "";
        string Results = "";
        int N = 0;
        int I = 0;
        if (IsInStr(Contents, "@NO-LINT-DFCTL")) return;
        vTypes = new List<dynamic>() { "CheckBox", "Command", "Option", "Frame", "Label", "TextBox", "RichTextBox", "RichTextBoxNew", "ComboBox", "ListBox", "Timer", "UpDown", "HScrollBar", "Image", "Picture", "MSFlexGrid", "DBGrid", "Line", "Shape", "DTPicker" };
        foreach (var itervType in vTypes)
        {
            vType = itervType;
            Matcher = "Begin [a-zA-Z0-9]*.[a-zA-Z0-9]* " + vType + "[0-9]*";
            N = RegExCount(Contents, Matcher);
            for (I = 0; I <= N - 1; I += 1)
            {
                Results = RegExNMatch(Contents, Matcher, I);
                RecordError(ref Errors, ref ErrorCount, TY_DFCTL, 0, "Default control name in use on form: " + Results);
            }
        }
    }
    public static void TestCodeLine(ref string Errors, ref int ErrorCount, int LineN, string L)
    {
        if (RegExTest(L, "+ \"") || RegExTest(L, "\" +")) RecordError(ref Errors, ref ErrorCount, TY_CORRE, LineN, "Possible use of + instead of & on String concatenation");
        if (RegExTest(L, " Me[.]")) RecordError(ref Errors, ref ErrorCount, TY_CORRE, LineN, "Use of 'Me.*' is not required.");
        if (RegExTest(L, "\\.Enabled = [-0-9]")) RecordError(ref Errors, ref ErrorCount, TY_CORRE, LineN, "Property [Enabled] Should Be Boolean.  Numeric found.");
        if (RegExTest(L, "\\.Visible = [-0-9]")) RecordError(ref Errors, ref ErrorCount, TY_CORRE, LineN, "Property [Visible] Should Be Boolean.  Numeric found.");
        if (RegExTest(L, " Call ")) RecordError(ref Errors, ref ErrorCount, TY_CORRE, LineN, "Remove keyword 'Call'.");
        if (RegExTest(L, " GoSub ") || RegExTest(L, " Return$")) RecordError(ref Errors, ref ErrorCount, TY_GOSUB, LineN, "Remove uses of 'GoSub' and 'Return'.");
        if (RegExTest(L, " Stop$") || RegExTest(L, " Return$")) RecordError(ref Errors, ref ErrorCount, TY_CSTOP, LineN, "Code contains STOP statement.");
    }
    public static void AddFix(string Typ, string Find, string Repl, bool RestOfFile = false)
    {
        int N = 0;
        if (InStr(AUTOFIX_LINT_TYPES, Typ) == 0) return;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (RestOfFile)
        {
            N = AutofixFindRestOfFile.Count;
            N = N + 1;
            // TODO: (NOT SUPPORTED): ReDim Preserve AutofixFindRestOfFile(1 To N)
            // TODO: (NOT SUPPORTED): ReDim Preserve AutofixReplRestOfFile(1 To N)
            AutofixFindRestOfFile[N] = Find;
            AutofixReplRestOfFile[N] = Repl;
        }
        else
        {
            N = AutofixFind.Count;
            N = N + 1;
            // TODO: (NOT SUPPORTED): ReDim Preserve AutofixFind(1 To N)
            // TODO: (NOT SUPPORTED): ReDim Preserve AutofixRepl(1 To N)
            AutofixFind[N] = Find;
            AutofixRepl[N] = Repl;
        }
    }
    public static int GetFixCount(bool RestOfFile = false)
    {
        int _GetFixCount = 0;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        _GetFixCount = (RestOfFile ? AutofixFindRestOfFile : AutofixFind).Count;
        return _GetFixCount;
    }
    public static string PerformAutofix(string Line)
    {
        string _PerformAutofix = "";
        int I = 0;
        int N = 0;
        string Find = "";
        string Repl = "";
        N = GetFixCount(false);
        if (N > 0)
        {
            for (I = 0; I <= AutofixFind.Count; I += 1)
            {
                Find = AutofixFind[I];
                Repl = AutofixRepl[I];
                if (Find == "") goto NextFix;
                Line = RegExReplace(Line, Find, Repl);
            NextFix:;
            }
        }
        N = GetFixCount(true);
        if (N > 0)
        {
            for (I = 0; I <= AutofixFindRestOfFile.Count; I += 1)
            {
                Find = AutofixFindRestOfFile[I];
                Repl = AutofixReplRestOfFile[I];
                if (Find == "") goto NextFixRestOfFile;
                Line = RegExReplace(Line, Find, Repl);
            NextFixRestOfFile:;
            }
        }
    Finish:;
        _PerformAutofix = Line;
        AutofixFind.Clear();
        AutofixRepl.Clear();
        return _PerformAutofix;
    }

}
