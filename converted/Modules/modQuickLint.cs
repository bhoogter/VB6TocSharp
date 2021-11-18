using ADODB;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Windows.Documents;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.DateAndTime;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modProjectFiles;
using static modRegEx;
using static VBExtension;


static class modQuickLint
{
    // Option Explicit
    private const int Idnt = 2;
    private const int MAX_ERRORS = 50;
    private const string Attr = "Attribute";
    private const string Q = "\"";
    private const string A = "'";
    private const string S = " ";
    private const string LintKey = "'@NO-LINT";
    private const string TY_ALLTY = "AllTy";
    private const string TY_ERROR = "Error";
    private const string TY_INDNT = "Indnt";
    private const string TY_ARGNA = "ArgNa";
    private const string TY_ARGTY = "ArgTy";
    private const string TY_FSPNA = "FSPNa";
    private const string TY_DEPRE = "Depre";
    private const string TY_MIGRA = "Migra";
    private const string TY_STYLE = "Style";
    private const string TY_BLANK = "Blank";
    private const string TY_EXPLI = "Expli";
    private const string TY_COMPA = "Compa";
    private const string TY_TYPEC = "TypeC";
    private const string TY_NOTYP = "NoTyp";
    private const string TY_BYRFV = "ByReV";
    private const string TY_PRIPU = "PriPu";
    private const string TY_FNCRE = "FncRe";
    private const string TY_CORRE = "Corre";
    private const string TY_GOSUB = "GoSub";
    private const string TY_CSTOP = "CStop";
    private const string TY_OPDEF = "OpDef";
    private const string TY_DFCTL = "DfCtl";
    public static string ErrorPrefix = "";
    public static string ErrorIgnore = "";


    public static dynamic ErrorTypes()
    {
        dynamic ErrorTypes = null;
        ErrorTypes = Array(TY_ALLTY, TY_ERROR, TY_INDNT, TY_ARGNA, TY_ARGTY, TY_FSPNA, TY_DEPRE, TY_MIGRA, TY_STYLE, TY_BLANK, TY_EXPLI, TY_COMPA, TY_TYPEC, TY_NOTYP, TY_BYRFV, TY_PRIPU, TY_FNCRE, TY_CORRE, TY_GOSUB, TY_CSTOP, TY_OPDEF, TY_DFCTL);
        return ErrorTypes;
    }

    public static string Lint(string FileName = "", bool Alert_UNUSED = true)
    {
        string Lint = "";
        string FileList = "";

        if (FileName == "")
        {
            FileName = "prj.vbp";
        }
        if (InStr(FileName, "\\") == 0)
        {
            FileName = AppDomain.CurrentDomain.BaseDirectory + "\\" + FileName;
        }
        FileList = IIf(Right(FileName, 4) == ".vbp", VBPCode(FileName), FileName);

        Lint = QuickLintFiles(FileList);
        return Lint;
    }

    public static string QuickLintFiles(string List_UNUSED)
    {
        string QuickLintFiles = "";
        const int lintDotsPerRow = 50;
        dynamic L = null;

        int x = 0;

        DateTime StartTime = DateTime.MinValue;

        StartTime = DateTime.Now; ;

        foreach (var iterL in Split(List, vbCrLf))
        {
            L = iterL;
            string Result = "";

            Result = QuickLintFile(L);
            if (!Result == "")
            {
                string S = "";

                Console.WriteLine(vbCrLf + "Done (" + DateDiff("s", StartTime, DateTime.Now;) +"s).  To re-run for failing file, hit enter on the line below:");
                S = "LINT FAILED: " + L + vbCrLf + Result + vbCrLf + "?Lint(\"" + L + "\")";
                QuickLintFiles = S;
                return QuickLintFiles;

            }
            else
            {
                Debug.PrintNNL(Switch(Right(L, 3) == "frm", "o", Right(L, 3) == "cls", "x", true, "."));
            }
            x = x + 1;
            if (x >= lintDotsPerRow)
            {
                x = 0;
                Console.WriteLine();
            }
            DoEvents();
        }
        Console.WriteLine(vbCrLf + "Done (" + DateDiff("s", StartTime, DateTime.Now;) +"s).");
        QuickLintFiles = "";
        return QuickLintFiles;
    }

    public static string QuickLintFile(string File)
    {
        string QuickLintFile = "";
        if (InStr(File, "\\") == 0)
        {
            File = AppDomain.CurrentDomain.BaseDirectory + "\\" + File;
        }
        string FName = "";
        string Contents = "";
        string GivenName = "";
        string CheckName = "";

        FName = Mid(File, InStrRev(File, "\\") + 1);
        CheckName = Replace(Replace(Replace(FName, ".bas", ""), ".cls", ""), ".frm", "");
        ErrorPrefix = Right(Space(18) + FName, 18) + " ";
        Contents = ReadEntireFile(File);
        GivenName = RegExNMatch(Contents, "Attribute VB_Name = \"([^\"]+)\"", 0);
        GivenName = Replace(Replace(GivenName, "Attribute VB_Name = ", ""), "\"", "");
        if (CheckName != GivenName)
        {
            QuickLintFile = "Module name [" + GivenName + "] must match file name [" + FName + "].  Rename module or class to match the other";
            return QuickLintFile;

        }
        QuickLintFile = QuickLintContents(Contents);
        return QuickLintFile;
    }

    public static string QuickLintContents(string Contents)
    {
        string QuickLintContents = "";
        List<string> Lines = new List<string> { }; // TODO - Specified Minimum Array Boundary Not Supported:   Dim Lines() As String, LL As Variant, L As String
        List<dynamic> LL = new List<dynamic> { }; // TODO - Specified Minimum Array Boundary Not Supported:   Dim Lines() As String, LL As Variant, L As String
        List<string> L = new List<string> { }; // TODO - Specified Minimum Array Boundary Not Supported:   Dim Lines() As String, LL As Variant, L As String

        // TODO (not supported): On Error GoTo LintError
        ErrorIgnore = "";
        Lines = Split(Replace(Contents, vbCr, ""), vbLf);

        bool InAttributes = false;
        bool InBody = false;


        string MultiLine = "";

        int Indent = 0;
        int LineN = 0;

        string Errors = "";
        int ErrorCount = 0;

        int BlankLineCount = 0;

        Collection Options = new Collection();


        Indent = 0;

        TestDefaultControlNames(Errors, ErrorCount, 0, Contents);


        foreach (var iterLL in Lines)
        {
            LL = iterLL;
            if (ErrorCount >= MAX_ERRORS)
            {
                break;
            }

            if (Right(LL, 2) == " _")
            {
                string Portion = "";

                Portion = Left(LL, Len(LL) - 2);
                if (MultiLine != "")
                {
                    Portion = Trim(Portion);
                }
                MultiLine = MultiLine + Portion;
                LineN = LineN + 1;
                goto NextLine;
            }
            else if (MultiLine != "")
            {
                LL = MultiLine + Trim(LL);
                MultiLine = "";
            }

            TestBlankLines(Errors, ErrorCount, LineN, LL, BlankLineCount);
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
            //If LineN = 15 Then Stop

            bool UnindentedAlready = false;

            if (RegExTest(L, "^Option "))
            {
                Options.Add("true", Replace(L, "Options ", ""));
            }
            else if (RegExTest(L, "^[ ]*(Else|ElseIf .* Then)$"))
            {
                Indent = Indent - Idnt;
            }
            else if (RegExTest(L, "^[ ]*End Select$"))
            {
                Indent = Indent - Idnt - Idnt;
            }
            else if (RegExTest(L, "^[ ]*(End (If|Function|Sub|Property|Enum|Type)|Next( .*)?|Wend|Loop|Loop (While .*|Until .*))$"))
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
            while (Mid(RTrim(L), LineIndent + 1, 1) == S)
            {
                LineIndent = LineIndent + 1;
            }
            TestIndent(Errors, ErrorCount, LineN, L, LineIndent, IIf(!RegExTest(L, "^[ ]*Case "), Indent, Indent - Idnt));

            List<string> Statements = new List<string> { }; // TODO - Specified Minimum Array Boundary Not Supported:     Dim Statements() As String, SS As Variant, St As String
            List<dynamic> SS = new List<dynamic> { }; // TODO - Specified Minimum Array Boundary Not Supported:     Dim Statements() As String, SS As Variant, St As String
            List<string> St = new List<string> { }; // TODO - Specified Minimum Array Boundary Not Supported:     Dim Statements() As String, SS As Variant, St As String

            Statements = Split(L, ": ");
            foreach (var iterSS in Statements)
            {
                SS = iterSS;
                St = SS;

                if (RegExTest(L, "^[ ]*(Else|ElseIf .* Then)$"))
                {
                    Indent = Indent + Idnt;
                }
                else if (RegExTest(St, "^[ ]*(End (If|Function|Sub|Property)|Next|Wend|Loop|Loop .*|Enum|Type|Select)$"))
                {
                    if (!UnindentedAlready)
                    {
                        Indent = Indent - Idnt;
                    }
                }
                else if (RegExTest(St, "^[ ]*If "))
                {
                    if (!RegExTest(St, "Then "))
                    {
                        Indent = Indent + Idnt;
                    }
                }
                else if (RegExTest(St, "^[ ]*For "))
                {
                    if (!RegExTest(St, " Next"))
                    {
                        Indent = Indent + Idnt;
                    }
                }
                else if (RegExTest(St, "^[ ]*Next$"))
                {
                    Indent = Indent - Idnt;
                }
                else if (RegExTest(St, "^[ ]*Next [a-zA-Z_][a-zA-Z0-9_]*$"))
                {
                    RecordError(Errors, ErrorCount, TY_STYLE, LineN, "Remove variable from NEXT statement");
                    Indent = Indent - Idnt;
                }
                else if (RegExTest(St, "^[ ]*While "))
                {
                    RecordError(Errors, ErrorCount, TY_STYLE, LineN, "Use Do While/Until...Loop in place of While...Wend");
                    if (!RegExTest(St, " Wend$"))
                    {
                        Indent = Indent + Idnt;
                    }
                }
                else if (RegExTest(St, "^[ ]*Do (While|Until)"))
                {
                    if (!RegExTest(St, ": Loop"))
                    {
                        Indent = Indent + Idnt;
                    }
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
                    RecordError(Errors, ErrorCount, TY_MIGRA, LineN, "Remove all uses of WITH.  No migration path exists.");
                }
                else if (RegExTest(St, "^[ ]*(Private |Public )?Declare (Function |Sub )"))
                {
                    // External Api
                }
                else if (RegExTest(St, "^((Private|Public|Friend) )?Function "))
                {
                    if (!RegExTest(St, ": End Function"))
                    {
                        Indent = Indent + Idnt;
                    }
                    TestSignature(Errors, ErrorCount, LineN, St);
                }
                else if (RegExTest(St, "^((Private|Public|Friend) )?Sub "))
                {
                    if (!RegExTest(St, ": End Sub"))
                    {
                        Indent = Indent + Idnt;
                    }
                    TestSignature(Errors, ErrorCount, LineN, St);
                }
                else if (RegExTest(St, "^((Private|Public|Friend) )?Property (Get|Let|Set) "))
                {
                    if (!RegExTest(St, ": End Property"))
                    {
                        Indent = Indent + Idnt;
                    }
                    TestSignature(Errors, ErrorCount, LineN, St);
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
                    TestDeclaration(Errors, ErrorCount, LineN, St, false);
                }
                else
                {
                    TestCodeLine(Errors, ErrorCount, LineN, St);
                }
            NextStatement:;
            }
        NextLine:;
        }

        TestModuleOptions(Errors, ErrorCount, Options);

        QuickLintContents = Errors;
        return QuickLintContents;

    LintError:;
        RecordError(Errors, ErrorCount, TY_ERROR, 0, "Linter Error [" + Err().Number + "]: " + Err().Description);
        QuickLintContents = Errors;
        return QuickLintContents;
    }

    private static string ReadEntireFile(string tFileName)
    {
        string ReadEntireFile = "";
        // TODO (not supported): On Error Resume Next
        dynamic mFSO = null;

        mFSO = CreateObject("Scripting.FileSystemObject");
        ReadEntireFile = mFSO.OpenTextFile(tFileName, 1).ReadAll;

        if (FileLen(tFileName()) / 10 != Len(ReadEntireFile) / 10)
        {
            MsgBox("ReadEntireFile was short: " + FileLen(tFileName()) + " vs " + Len(ReadEntireFile));
        }
        return ReadEntireFile;
    }

    public static string CleanLine(string Line)
    {
        string CleanLine = "";
        int x = 0;
        int Y = 0;

        while (true)
        {
            x = InStr(Line, Q);
            if (x == 0)
            {
                break;
            }

            Y = InStr(x + 1, Line, Q);
            while (Mid(Line, Y + 1, 1) == Q)
            {
                Y = InStr(Y + 2, Line, Q);
            }

            if (Y == 0)
            {
                break;
            }
            Line = Left(Line, x - 1) + String(Y - x + 1, "S") + Mid(Line, Y + 1);
        }

        x = InStr(Line, A);
        if (x > 0)
        {
            Line = RTrim(Left(Line, x - 1));
        }

        CleanLine = Line;
        return CleanLine;
    }

    public static void RecordError(ref string Errors, out int ErrorCount, string Typ, int LineN, string Error)
    {
        if (InStr(ErrorIgnore, UCase(Typ)) > 0 || InStr(ErrorIgnore, TY_ALLTY) > 0)
        {
            return;

        }

        if (Len(Errors) != 0)
        {
            Errors = Errors + vbCrLf;
        }
        if (InStr(Join(ErrorTypes(), ","), Typ) == 0)
        {
            Errors = Errors + ErrorPrefix + "[" + TY_ERROR + "] Line " + Right(Space(5) + LineN, 5) + ": Unknown error type in linter (add to ErrorTypes): " + Typ;
        }
        Errors = Errors + ErrorPrefix + "[" + Right(Space(5) + Typ, 5) + "] Line " + Right(Space(5) + LineN, 5) + ": " + Error;
        ErrorCount = ErrorCount + 1;
    }

    public static bool StartsWith(string L, string Find)
    {
        bool StartsWith = false;
        StartsWith = Left(L, Len(Find)) == Find;
        return StartsWith;
    }

    public static string StripLeft(string L, string Find)
    {
        string StripLeft = "";
        if (StartsWith(L, Find))
        {
            StripLeft = Mid(L, Len(Find) + 1);
        }
        else
        {
            StripLeft = L;
        }
        return StripLeft;
    }

    public static void TestIndent(ref string Errors, ref int ErrorCount, int LineN, string L, int LineIndent, int ExpectedIndent)
    {
        if (RTrim(L) == "")
        {
            return;

        }
        if (RegExTest(L, "^On Error "))
        {
            return;

        }
        if (RegExTest(L, "^[a-zA-Z][a-zA-Z0-9]*:$"))
        {
            return;

        }

        if (LineIndent != ExpectedIndent)
        {
            RecordError(Errors, ErrorCount, TY_INDNT, LineN, "Incorrect Indent -- expected " + ExpectedIndent + ", got " + LineIndent);
        }
    }

    public static void TestBlankLines(ref string Errors, ref int ErrorCount, int LineN, string L, out int BlankLineCount)
    {
        if (Trim(L) != "")
        {
            BlankLineCount = 0;
            return;

        }
        BlankLineCount = BlankLineCount + 1;
        if (BlankLineCount > 3)
        {
            RecordError(Errors, ErrorCount, TY_BLANK, LineN, "Too many blank lines.");
        }
    }

    public static void TestLintControl(string L)
    {
        dynamic LL = null;

        if (InStr(L, LintKey) == 0)
        {
            return;

        }
        string Match = "";
        string Typ = "";

        Match = RegExNMatch(L, LintKey + "(-.....)?", 0);
        Typ = IIf(Match == LintKey, TY_ALLTY, Replace(Match, LintKey + "-", ""));
        ErrorIgnore = ErrorIgnore + "," + Typ;
    }

    public static void TestModuleOptions(ref string Errors, ref int ErrorCount, Collection Options)
    {
        // TODO (not supported): On Error Resume Next
        string Value = "";

        Value = "";
        Value = Options["Explicit"];
        if (Value != "")
        {
            RecordError(Errors, ErrorCount, TY_EXPLI, 0, "Option Explicit not set on file");
        }

        Value = "";
        Value = Options["Compare Binary"];
        Value = Options["Compare Database"];
        if (Value != "")
        {
            RecordError(Errors, ErrorCount, TY_COMPA, 0, "Use of Option Compare not recommended");
        }
    }

    public static void TestArgName(ref string Errors, ref int ErrorCount, int LineN, string Name)
    {
        string LL = "";

        LL = Trim(Name);

        if (RegExTest(LL, "^[a-z][a-z0-9_]*$"))
        {
            RecordError(Errors, ErrorCount, TY_ARGNA, LineN, "Identifier name declared as all lower-case: " + LL);
        }

        if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*%$"))
        { // % Integer Dim L%
            RecordError(Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Integer deprecated: " + LL);
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*&$"))
        { // & Long  Dim M&
            RecordError(Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Long deprecated: " + LL);
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*@$"))
        { // @ Decimal Const W@ = 37.5
            RecordError(Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Decimal deprecated: " + LL);
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-TY_TYPEC-Z0-9_]*!$"))
        { // ! Single  Dim Q!
            RecordError(Errors, ErrorCount, TY_DEPRE, LineN, "Use of Type Character For Single deprecated: " + LL);
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*#$"))
        { // # Double  Dim X#
            RecordError(Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For Double deprecated: " + LL);
        }
        else if (RegExTest(LL, "^[a-zA-Z_][a-zA-Z0-9_]*\\$$"))
        { // $ String  Dim V$ = "Secret"
            RecordError(Errors, ErrorCount, TY_TYPEC, LineN, "Use of Type Character For String deprecated: " + LL);
        }
    }

    public static void TestSignatureName(ref string Errors, ref int ErrorCount, int LineN, string Name)
    {
        string LL = "";

        LL = Trim(Name);

        if (RegExTest(LL, "^[a-z][a-z0-9_]*$"))
        {
            RecordError(Errors, ErrorCount, TY_FSPNA, LineN, "Func/Sub/Prop name declared as all lower-case: " + LL);
        }
    }

    public static void TestDeclaration(ref string Errors, ref int ErrorCount, int LineN, string L, bool InSignature)
    {
        bool IsOptional = false;
        bool IsByVal = false;
        bool IsByRef = false;
        bool IsParamArray = false;

        L = Trim(L);
        L = StripLeft(L, "Dim ");
        L = StripLeft(L, "Private ");
        L = StripLeft(L, "Public ");
        L = StripLeft(L, "Const ");
        L = StripLeft(L, "Global ");

        dynamic LL = null;

        foreach (var iterLL in Split(L, ", "))
        {
            LL = iterLL;
            int Ix = 0;
            string ArgName = "";
            string ArgType = "";
            string ArgDefault = "";


            IsOptional = StartsWith(LL, "Optional ");
            LL = StripLeft(LL, "Optional ");

            IsByVal = StartsWith(LL, "ByVal ");
            LL = StripLeft(LL, "ByVal ");

            IsByRef = StartsWith(LL, "ByRef ");
            LL = StripLeft(LL, "ByRef ");

            IsParamArray = StartsWith(LL, "ParamArray ");
            LL = StripLeft(LL, "ParamArray ");

            Ix = InStr(LL, " = ");
            if (Ix > 0)
            {
                ArgDefault = Trim(Mid(LL, Ix + 3));
                LL = Left(LL, Ix - 1);
            }
            else
            {
                ArgDefault = "";
            }

            Ix = InStr(LL, " As ");
            if (Ix > 0)
            {
                ArgType = Trim(Mid(LL, Ix + 4));
                LL = Left(LL, Ix - 1);
            }
            else
            {
                ArgType = "";
            }

            //    If IsParamArray Then Stop
            if (ArgType == "")
            {
                RecordError(Errors, ErrorCount, TY_NOTYP, LineN, "Local Parameter Missing Type: [" + LL + "]");
            }
            if (InSignature)
            {
                if (IsParamArray)
                {
                    if (Right(LL, 2) != "()")
                    {
                        RecordError(Errors, ErrorCount, TY_STYLE, LineN, "ParamArray variable not declared as an Array.  Add '()': " + LL);
                    }
                }
                else
                {
                    if (!IsByVal && !IsByRef)
                    {
                        RecordError(Errors, ErrorCount, TY_BYRFV, LineN, "ByVal or ByRef not specified on parameter [" + LL + "] -- specify one or the other");
                    }
                }
                if (IsOptional && ArgDefault == "")
                {
                    RecordError(Errors, ErrorCount, TY_OPDEF, LineN, "Parameter declared OPTIONAL but no default specified. Must specify default: " + LL);
                }
            }

            TestArgName(Errors, ErrorCount, LineN, LL);

            TestArgType(Errors, ErrorCount, LineN, LL, ArgType);
        }
    }

    public static void TestArgType(ref string Errors, ref int ErrorCount, int LineN, string Name, string Typ)
    {
        if (Typ == "Integer")
        {
            RecordError(Errors, ErrorCount, TY_ARGTY, LineN, "Arg [" + Name + "] is of type [" + Typ + "] -- use Long");
        }
        if (Typ == "Short")
        {
            RecordError(Errors, ErrorCount, TY_ARGTY, LineN, "Arg [" + Name + "] is of type [" + Typ + "] -- use Long");
        }
        if (Typ == "Byte")
        {
            RecordError(Errors, ErrorCount, TY_ARGTY, LineN, "Arg [" + Name + "] is of type [" + Typ + "] -- use Long");
        }
        if (Typ == "Float")
        {
            RecordError(Errors, ErrorCount, TY_ARGTY, LineN, "Arg [" + Name + "] is of type [" + Typ + "] -- use Double");
        }
    }

    public static void TestSignature(ref string Errors, ref int ErrorCount, int LineN, string LL)
    {
        if (!RegExTest(LL, "^[ ]*(Private|Public|Friend) "))
        {
            RecordError(Errors, ErrorCount, TY_PRIPU, LineN, "Either Private or Public should be specified, but neither was.");
        }

        string L = "";
        bool WithReturn = false;

        L = LL;
        L = StripLeft(L, "Private ");
        L = StripLeft(L, "Public ");
        L = StripLeft(L, "Friend ");
        L = StripLeft(L, "Sub ");
        if (StartsWith(L, "Function ") || StartsWith(L, "Property Get "))
        {
            WithReturn = true;
        }
        L = StripLeft(L, "Function ");
        L = StripLeft(L, "Property ");

        int Ix = 0;
        int Ix2 = 0;
        string Name = "";
        string Args = "";
        string Ret = "";

        Ix = InStr(L, "(");
        if (Ix == 0)
        {
            return;

        }
        Name = Left(L, Ix - 1);
        if (RegExTest(L, "\\) As .*\\(\\)"))
        {
            Ix2 = InStrRev(L, ")", Len(L) - 2);
        }
        else
        {
            Ix2 = InStrRev(L, ")");
        }
        Args = Mid(L, Ix + 1, Ix2 - Ix - 1);
        Ret = Mid(L, Ix2 + 1);

        TestSignatureName(Errors, ErrorCount, LineN, Name);
        if (WithReturn && Ret == "")
        {
            RecordError(Errors, ErrorCount, TY_FNCRE, LineN, "Function Return Type Not Specified -- Specify Return Type or Variant");
        }
        TestDeclaration(Errors, ErrorCount, LineN, Args, true);
    }

    public static void TestDefaultControlNames(ref string Errors, ref int ErrorCount, int LineN_UNUSED, string Contents)
    {
        List<dynamic> vTypes = new List<dynamic> { }; // TODO - Specified Minimum Array Boundary Not Supported:   Dim vTypes() As Variant, vType As Variant
        List<dynamic> vType = new List<dynamic> { }; // TODO - Specified Minimum Array Boundary Not Supported:   Dim vTypes() As Variant, vType As Variant

        string Matcher = "";
        string Results = "";
        int N = 0;
        int I = 0;

        vTypes = Array("CheckBox", "Command", "Option", "Frame", "Label", "TextBox", "RichTextBox", "RichTextBoxNew", "ComboBox", "ListBox", "Timer", "UpDown", "HScrollBar", "Image", "Picture", "MSFlexGrid", "DBGrid", "Line", "Shape", "DTPicker");

        foreach (var itervType in vTypes)
        {
            vType = itervType;
            Matcher = "Begin [a-zA-Z0-9]*.[a-zA-Z0-9]* " + vType + "[0-9]*";
            N = RegExCount(Contents, Matcher);
            for (I = 0; I < N - 1; I++)
            {
                Results = RegExNMatch(Contents, Matcher, I);
                RecordError(Errors, ErrorCount, TY_DFCTL, 0, "Default control name in use on form: " + Results);
            }
        }
    }

    public static void TestCodeLine(ref string Errors_UNUSED, ref int ErrorCount, int LineN, string L)
    {
        if (RegExTest(L, "+ \"") || RegExTest(L, "\" +"))
        {
            RecordError(Error, ErrorCount, TY_CORRE, LineN, "Possible use of + instead of & on String concatenation");
        }
        if (RegExTest(L, " Me[.]"))
        {
            RecordError(Error, ErrorCount, TY_CORRE, LineN, "Use of 'Me.*' is not required.");
        }

        if (RegExTest(L, "\\.Enabled = [-0-9]"))
        {
            RecordError(Error, ErrorCount, TY_CORRE, LineN, "Property [Enabled] Should Be Boolean.  Numeric found.");
        }
        if (RegExTest(L, "\\.Visible = [-0-9]"))
        {
            RecordError(Error, ErrorCount, TY_CORRE, LineN, "Property [Visible] Should Be Boolean.  Numeric found.");
        }

        if (RegExTest(L, " Call "))
        {
            RecordError(Error, ErrorCount, TY_CORRE, LineN, "Remove keyword 'Call'.");
        }
        if (RegExTest(L, " GoSub ") || RegExTest(L, " Return$"))
        {
            RecordError(Error, ErrorCount, TY_GOSUB, LineN, "Remove uses of 'GoSub' and 'Return'.");
        }

        if (RegExTest(L, " Stop$") || RegExTest(L, " Return$"))
        {
            RecordError(Error, ErrorCount, TY_CSTOP, LineN, "Code contains STOP statement.");
        }
    }
}
