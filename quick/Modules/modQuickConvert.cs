using System;
using System.Collections.Generic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Conversion;
using static Microsoft.VisualBasic.DateAndTime;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static Microsoft.VisualBasic.VBMath;
using static modControlProperties;
using static modConvertForm;
using static modProjectFiles;
using static modQuickLint;
using static modRefScan;
using static modRegEx;
using static modUtils;
using static VBExtension;



static class modQuickConvert
{
    public const int Idnt = 2;
    public const string Attr = "Attribute";
    public const string Q = "\"";
    public const string A = "'";
    public const string S = " ";
    public const string STRING_TOKEN_PREFIX = "__S";
    public const string EXPRESSION_TOKEN_PREFIX = "__E_";
    public static List<string> LineStrings = new List<string>();
    public static int LineStringsCount = 0;
    public static string LineComment = "";
    public static bool InProperty = false;
    public static string CurrentTypeName = "";
    public static string CurrentEnumName = "";
    public static string CurrentFunctionName = "";
    public static string CurrentFunctionReturnValue = "";
    public static string CurrentFunctionArgs = "";
    public static string CurrentFunctionArrays = "";
    public static string ModuleName = "";
    public static string ModuleFunctions = "";
    public static string ModuleArrays = "";
    public static string ModuleProperties = "";
    public enum DeclarationType
    {
        DECL_GLOBAL = 99
    , DECL_SIGNATURE = 98
    , DECL_LOCAL = 1
    , DECL_TYPE
    , DECL_ENUM
    , DECL_EXTERN = 101
    }
    public enum CodeType
    {
        CODE_MODULE
    , CODE_CLASS
    , CODE_FORM
    , CODE_CONTROL
    }
    public class RandomType
    {
        public int J;
        public string W;
        public string X; // TODO: (NOT SUPPORTED) Fixed Length String not supported: (5)
    }
    private static string ResolveSources(string FileName)
    {
        string _ResolveSources = "";
        if (FileName == "") FileName = "prj.vbp";
        if (FileName == "forms")
        {
            _ResolveSources = VBPForms("true");
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
    public static string Convert(string FileName = "")
    {
        string _Convert = "";
        string FileList = "";
        FileList = ResolveSources(FileName);
        _Convert = QuickConvertFiles(FileList);
        return _Convert;
    }
    public static string QuickConvertFiles(string List)
    {
        string _QuickConvertFiles = "";
        int lintDotsPerRow = 50;
        dynamic L = null;
        int X = 0;
        DateTime StartTime = DateTime.MinValue;
        StartTime = DateTime.Now;
        foreach (var iterL in new List<string>(Split(List, vbCrLf)))
        {
            L = iterL;
            string Result = "";
            Result = QuickConvertFile(L);
            if (Result != "")
            {
                string S = "";
                Console.WriteLine(vbCrLf + "Done (" + DateDiff("s", StartTime, DateTime.Now) + "s).  To re-run for failing file, hit enter on the line below:");
                S = "LINT FAILED: " + L + vbCrLf + Result + vbCrLf + "?Lint(\"" + L + "\")";
                _QuickConvertFiles = S;
                return _QuickConvertFiles;
            }
            else
            {
                //Debug.Print(Switch(Right(L, 3) == "frm", "o", Right(L, 3) == "cls", "x", true, "."));
            }
            X = X + 1;
            if (X >= lintDotsPerRow) { X = 0; }
            DoEvents();
        }
        Console.WriteLine(vbCrLf + "Done (" + DateDiff("s", StartTime, DateTime.Now) + "s).");
        _QuickConvertFiles = "";
        return _QuickConvertFiles;
    }
    public static CodeType CodeFileType(string File)
    {
        CodeType _CodeFileType = CodeType.CODE_MODULE;
        switch (Right(LCase(File), 4))
        {
            case ".bas":
                _CodeFileType = CodeType.CODE_MODULE;
                break;
            case ".frm":
                _CodeFileType = CodeType.CODE_FORM;
                break;
            case ".cls":
                _CodeFileType = CodeType.CODE_CLASS;
                break;
            case ".ctl":
                _CodeFileType = CodeType.CODE_CONTROL;
                break;
            default:
                _CodeFileType = CodeType.CODE_MODULE;
                break;
        }
        return _CodeFileType;
    }
    public static string QuickConvertFile(string File)
    {
        string _QuickConvertFile = "";
        ModuleArrays = "";
        if (InStr(File, "\\") == 0) File = AppContext.BaseDirectory + "\\" + File;
        string fName = "";
        string Contents = "";
        string GivenName = "";
        string CheckName = "";
        fName = Mid(File, InStrRev(File, "\\") + 1);
        CheckName = Replace(Replace(Replace(fName, ".bas", ""), ".cls", ""), ".frm", "");
        ErrorPrefix = Right(Space(18) + fName, 18) + " ";
        Contents = ReadEntireFile(File);
        GivenName = GetModuleName(Contents);
        if (LCase(CheckName) != LCase(GivenName))
        {
            _QuickConvertFile = "Module name [" + GivenName + "] must match file name [" + fName + "].  Rename module or class to match the other";
            return _QuickConvertFile;
        }
        _QuickConvertFile = ConvertContents(Contents, CodeFileType(File));
        return _QuickConvertFile;
    }
    public static string GetModuleName(string Contents)
    {
        string _GetModuleName = "";
        _GetModuleName = RegExNMatch(Contents, "Attribute VB_Name = \"([^\"]+)\"", 0);
        _GetModuleName = Replace(Replace(_GetModuleName, "Attribute VB_Name = ", ""), "\"", "");
        return _GetModuleName;
    }
    public static string I(int N)
    {
        string _I = "";
        if (N <= 0) _I = ""; _I = Space(N);
        return _I;
    }
    public static string ConvertContents(string Contents, CodeType vCodeType, bool SubSegment = false)
    {
        string _ConvertContents = "";
        List<string> Lines = new List<string>();
        dynamic ActualLine = null;
        string LL = "";
        string L = "";
        // On Error GoTo LintError
        if (!SubSegment)
        {
            ModuleName = GetModuleName(Contents);
            ModuleFunctions = GetModuleFunctions(Contents);
        }
        Lines = new List<string>(Split(Replace(Contents, vbCr, ""), vbLf));
        bool InAttributes = false;
        bool InBody = false;
        InBody = SubSegment;
        string MultiLineOrig = "";
        string MultiLine = "";
        bool IsMultiLine = false;
        int LineN = 0;
        int Indent = 0;
        string NewContents = "";
        bool SelectHasCase = false;
        Indent = 0;
        NewContents = "";
        // NewContents = UsingEverything & vbCrLf2
        // NewContents = NewContents & __S1 & ModuleName & __S2 & vbCrLf
        foreach (var iterActualLine in Lines)
        {
            ActualLine = iterActualLine;
            LL = ActualLine;
            // If MaxErrors > 0 And ErrorCount >= MaxErrors Then Exit For
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
            L = CleanLine(LL);
            if (!InBody)
            {
                bool IsAttribute = false;
                IsAttribute = StartsWith(LTrim(L), "Attribute ");
                if (!InAttributes && IsAttribute)
                {
                    InAttributes = true;
                    goto NextLineWithoutRecord;
                }
                else if (InAttributes && !IsAttribute)
                {
                    InAttributes = false;
                    InBody = true;
                    LineN = 0;
                }
                else
                {
                    goto NextLineWithoutRecord;
                }
            }
            LineN = LineN + 1;
            // If LineN >= 8 Then Stop
            bool UnindentedAlready = false;
            if (RegExTest(L, "^[ ]*(Else|ElseIf .* Then)$"))
            {
                Indent = Indent - Idnt;
                UnindentedAlready = true;
            }
            else if (RegExTest(L, "^[ ]*End Select$"))
            {
                Indent = Indent - Idnt - Idnt;
            }
            else if (RegExTest(L, "^[ ]*(End (If|Function|Sub|Property|Enum|Type)|Next( .*)?|Wend|Loop|Loop (While .*|Until .*)|ElseIf .*)$"))
            {
                Indent = Indent - Idnt;
                UnindentedAlready = true;
                CurrentEnumName = "";
                CurrentTypeName = "";
            }
            else
            {
                UnindentedAlready = false;
            }
            string NewLine = "";
            NewLine = "";
            if (InProperty)
            { // we process properties out of band to keep getters and setters together
                if (InStr(L, "End Property") > 0) InProperty = false;
                goto NextLineWithoutRecord;
            }
            if (CurrentTypeName != "")
            { // if we are in a type or an enum, the entire line is parsed as such
                NewLine = NewLine + ConvertTypeLine(L, vCodeType);
            }
            else if (CurrentEnumName != "")
            {
                NewLine = NewLine + ConvertEnumLine(L);
            }
            else if (RegExTest(L, "^[ ]*If "))
            { // The __S2 control structure, when single-line, lacks the __S3 to signal a close.
                NewLine = NewLine + ConvertIf(L);
                if (InStr(L, " Then ") == 0) Indent = Indent + Idnt;
            }
            else if (RegExTest(L, "^[ ]*ElseIf .*$"))
            {
                NewLine = NewLine + ConvertIf(L);
                if (InStr(L, " Then ") == 0) Indent = Indent + Idnt;
            }
            else
            {
                List<string> Statements = new List<string>();
                dynamic SS = null;
                string St = "";
                Statements = new List<string>(Split(Trim(L), ": "));
                foreach (var iterSS in Statements)
                {
                    SS = iterSS;
                    St = SS;
                    if (RegExTest(St, "^[ ]*ElseIf .*$"))
                    {
                        NewLine = NewLine + ConvertIf(St);
                        Indent = Indent + Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*Else$"))
                    {
                        NewLine = NewLine + "} else {";
                        Indent = Indent + Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*End Function"))
                    {
                        NewLine = NewLine + "return " + CurrentFunctionReturnValue + ";" + vbCrLf + "}";
                        CurrentFunctionName = "";
                        CurrentFunctionReturnValue = "";
                        CurrentFunctionArrays = "";
                        if (!UnindentedAlready) Indent = Indent - Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*End Select$"))
                    {
                        NewLine = NewLine + "break;" + vbCrLf;
                        NewLine = NewLine + "}";
                        if (!UnindentedAlready) Indent = Indent - Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*End (If|Sub|Enum|Type)$"))
                    {
                        CurrentTypeName = "";
                        CurrentEnumName = "";
                        NewLine = NewLine + "}";
                        if (!UnindentedAlready) Indent = Indent - Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*For Each"))
                    {
                        Indent = Indent + Idnt;
                        NewLine = ConvertForEach(St);
                    }
                    else if (RegExTest(St, "^[ ]*For "))
                    {
                        Indent = Indent + Idnt;
                        NewLine = ConvertFor(St);
                    }
                    else if (RegExTest(St, "^[ ]*Next\\b"))
                    {
                        NewLine = NewLine + "}";
                        if (!UnindentedAlready) Indent = Indent - Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*While "))
                    {
                        NewLine = NewLine + ConvertWhile(St);
                        Indent = Indent + Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*Wend"))
                    {
                        NewLine = NewLine + "}";
                        if (!UnindentedAlready) Indent = Indent - Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*Do (While|Until)"))
                    {
                        NewLine = NewLine + ConvertWhile(St);
                        Indent = Indent + Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*Loop$"))
                    {
                        NewLine = NewLine + "}";
                    }
                    else if (RegExTest(St, "^[ ]*Do$"))
                    {
                        NewLine = NewLine + "do {";
                        Indent = Indent + Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*(Loop While |Loop Until )"))
                    {
                        NewLine = NewLine + ConvertWhile(St);
                    }
                    else if (RegExTest(St, "^[ ]*Select Case "))
                    {
                        NewLine = NewLine + ConvertSwitch(St);
                        Indent = Indent + Idnt + Idnt;
                        SelectHasCase = false;
                    }
                    else if (RegExTest(St, "^[ ]*Case "))
                    {
                        NewLine = NewLine + ConvertSwitchCase(St, SelectHasCase);
                        SelectHasCase = true;
                    }
                    else if (RegExTest(St, "^[ ]*(Private |Public )?Declare (Function |Sub )"))
                    {
                        NewLine = NewLine + ConvertDeclare(St); // External Api
                    }
                    else if (RegExTest(St, "^((Private|Public|Friend) )?Function "))
                    {
                        CurrentFunctionArgs = "";
                        Indent = Indent + Idnt;
                        NewLine = NewLine + ConvertSignature(St, vCodeType);
                    }
                    else if (RegExTest(St, "^((Private|Public|Friend) )?Sub "))
                    {
                        CurrentFunctionArgs = "";
                        Indent = Indent + Idnt;
                        NewLine = NewLine + ConvertSignature(St, vCodeType);
                    }
                    else if (RegExTest(St, "^((Private|Public|Friend) )?Property (Get|Let|Set) "))
                    {
                        CurrentFunctionArgs = "";
                        NewLine = NewLine + ConvertProperty(St, Contents, vCodeType);
                        InProperty = true;
                        Indent = Indent + Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*(Public |Private )?Enum "))
                    {
                        NewLine = NewLine + ConvertEnum(St);
                        Indent = Indent + Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*(Public |Private )?Type "))
                    {
                        NewLine = NewLine + ConvertType(St);
                        Indent = Indent + Idnt;
                    }
                    else if (RegExTest(St, "^[ ]*(Dim|Private|Public|Const|Global|Static) "))
                    {
                        NewLine = NewLine + ConvertDeclaration(St, CurrentFunctionName == "" ? DeclarationType.DECL_GLOBAL : DeclarationType.DECL_LOCAL, vCodeType);
                    }
                    else
                    {
                        NewLine = NewLine + ConvertStatement(St);
                    }
                NextStatement:;
                }
            }
        NextLine:;
            // If IsMultiLine Then Stop
            // If InStr(LL, __S1) > 0 Then Stop
            // If InStr(LL, __S1) > 0 Then Stop
            // If Indent < 0 Then Stop
            NewLine = Decorate(NewLine);
            if (Trim(NewLine) != "")
            {
                NewContents = NewContents + I(Indent) + NewLine + vbCrLf;
            }
        NextLineWithoutRecord:;
        }
        // If AutoFix <> __S1 Then WriteFile AutoFix, Left(NewContents, Len(NewContents) - 2), True
        // NewContents = NewContents & __S1 & vbCrLf
        _ConvertContents = NewContents;
        return _ConvertContents;
    LintError:;
        Console.WriteLine("Error in quick convert [" + Err().Number + "]: " + Err().Description);
        _ConvertContents = "Error in quick convert [" + Err().Number + "]: " + Err().Description;
        return _ConvertContents;
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
        string Token = "";
        string Value = "";
        LineStrings.Clear();
        LineStringsCount = 0;
        LineComment = "";
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
            LineStringsCount = LineStringsCount + 1;
            // TODO: (NOT SUPPORTED): ReDim Preserve LineStrings(1 To LineStringsCount)
            Value = ConvertStringLiteral(Mid(Line, X, Y - X + 1));
            LineStrings[LineStringsCount] = Value;
            Token = STRING_TOKEN_PREFIX + LineStringsCount;
            Line = Left(Line, X - 1) + Token + Mid(Line, Y + 1);
        }
        X = InStr(Line, A);
        if (X > 0)
        {
            LineComment = Trim(Mid(Line, X + 1));
            Line = RTrim(Left(Line, X - 1));
        }
        _CleanLine = Line;
        return _CleanLine;
    }
    public static string Decorate(string Line)
    {
        string _Decorate = "";
        int I = 0;
        for (I = LineStringsCount; I <= -1; I += 1)
        {
            Line = Replace(Line, "__S" + I, LineStrings[I]);
        }
        if (LineComment != "") Line = Line + " // " + LineComment;
        _Decorate = Line;
        return _Decorate;
    }
    public static string ConvertStringLiteral(string L)
    {
        string _ConvertStringLiteral = "";
        L = Replace(L, "\\", "\\\\");
        L = "\"" + Replace(Mid(L, 2, Len(L) - 2), "\"\"", "\\\"") + "\"";
        _ConvertStringLiteral = L;
        return _ConvertStringLiteral;
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
        if (StartsWith(L, Find)) _StripLeft = Mid(L, Len(Find) + 1); _StripLeft = L;
        return _StripLeft;
    }
    public static bool RecordLeft(ref string L, string Find)
    {
        bool _RecordLeft = false;
        _RecordLeft = StartsWith(L, Find);
        if (_RecordLeft) L = Mid(L, Len(Find) + 1);
        return _RecordLeft;
    }
    public static string RemoveUntil(ref string L, string Find, bool RemoveFind = false)
    {
        string _RemoveUntil = "";
        int IX = 0;
        IX = InStr(L, Find);
        if (IX <= 0) return _RemoveUntil;
        _RemoveUntil = Left(L, IX - 1);
        L = Mid(L, IIf(RemoveFind, IX + Len(Find), IX));
        return _RemoveUntil;
    }
    private static string GetModuleFunctions(string Contents)
    {
        string _GetModuleFunctions = "";
        string Pattern = "(Private (Function|Sub) [^(]+\\()";
        int N = 0;
        int I = 0;
        string S = "";
        N = RegExCount(Contents, Pattern);
        _GetModuleFunctions = "";
        for (I = 0; I <= N - 1; I += 1)
        {
            S = RegExNMatch(Contents, Pattern, I);
            S = Replace(S, "Private ", "");
            S = Replace(S, "Sub ", "");
            S = Replace(S, "Function ", "");
            S = Replace(S, "(", "");
            _GetModuleFunctions = _GetModuleFunctions + "[" + S + "]";
        }
        return _GetModuleFunctions;
    }
    private static bool IsLocalFuncRef(string F)
    {
        bool _IsLocalFuncRef = false;
        _IsLocalFuncRef = InStr(ModuleFunctions, "[" + Trim(F) + "]") != 0;
        return _IsLocalFuncRef;
    }
    private static int SearchLeft(int Start, string Src, string Find, bool NotIn = false, bool Reverse = false)
    {
        int _SearchLeft = 0;
        int Bg = 0;
        int Ed = 0;
        int St = 0;
        int I = 0;
        string C = "";
        bool Found = false;
        if (!Reverse)
        {
            Bg = (Start == 0 ? 1 : Start);
            Ed = Len(Src);
            St = 1;
        }
        else
        {
            Bg = (Start == 0 ? Len(Src) : Start);
            Ed = 1;
            St = -1;
        }
        for (I = Bg; I <= St; I += Ed)
        {
            C = Mid(Src, I, 1);
            Found = InStr(Find, C) > 0;
            if (!NotIn && Found || NotIn && !Found)
            {
                _SearchLeft = I;
                return _SearchLeft;
            }
        }
        _SearchLeft = 0;
        return _SearchLeft;
    }
    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    public static string ConvertIf(string L)
    {
        string _ConvertIf = "";
        int ixThen = 0;
        string Expression = "";
        bool WithThen = false;
        bool WithElse = false;
        bool MultiStatement = false;
        L = Trim(L);
        ixThen = InStr(L, " Then");
        WithThen = InStr(L, " Then ") > 0;
        WithElse = InStr(L, " Else ") > 0;
        Expression = Trim(Left(L, ixThen - 1));
        Expression = StripLeft(Expression, "If ");
        Expression = StripLeft(Expression, "ElseIf ");
        _ConvertIf = InStr(L, "ElseIf") == 0 ? "if" : "} else if";
        _ConvertIf = _ConvertIf + "(" + ConvertExpression(Expression) + ")";
        if (!WithThen)
        {
            _ConvertIf = _ConvertIf + " {";
        }
        else
        {
            string cThen = "";
            string cElse = "";
            cThen = Trim(Mid(L, ixThen + 5));
            int ixElse = 0;
            ixElse = InStr(cThen, " Else ");
            if (ixElse > 0)
            {
                cElse = Mid(cThen, ixElse + 6);
                cThen = Left(cThen, ixElse - 1);
            }
            else
            {
                cElse = "";
            }
            // Inline Then
            dynamic St = null;
            MultiStatement = InStr(cThen, ": ") > 0;
            if (MultiStatement)
            {
                _ConvertIf = _ConvertIf + " { ";
                foreach (var iterSt in new List<string>(Split(cThen, ": ")))
                {
                    St = iterSt;
                    _ConvertIf = _ConvertIf + ConvertStatement(St) + " ";
                }
                _ConvertIf = _ConvertIf + "}";
            }
            else
            {
                _ConvertIf = _ConvertIf + ConvertStatement(cThen);
            }
            // Inline Then ... Else
            if (ixElse > 0)
            {
                MultiStatement = InStr(cElse, ":") > 0;
                if (MultiStatement)
                {
                    _ConvertIf = _ConvertIf + " { ";
                    foreach (var iterSt in new List<string>(Split(cElse, ":")))
                    {
                        St = iterSt;
                        _ConvertIf = _ConvertIf + ConvertStatement(Trim(St));
                    }
                    _ConvertIf = _ConvertIf + " }";
                }
                else
                {
                    _ConvertIf = _ConvertIf + ConvertStatement(cElse);
                }
            }
        }
        return _ConvertIf;
    }
    public static string ConvertSwitch(string L)
    {
        string _ConvertSwitch = "";
        _ConvertSwitch = "switch(" + ConvertExpression(Trim(Replace(L, "Select Case ", ""))) + ") {";
        return _ConvertSwitch;
    }
    public static string ConvertSwitchCase(string L, bool SelectHasCase)
    {
        string _ConvertSwitchCase = "";
        dynamic V = null;
        _ConvertSwitchCase = "";
        if (SelectHasCase) _ConvertSwitchCase = _ConvertSwitchCase + "break;" + vbCrLf;
        if (Trim(L) == "Case Else")
        {
            _ConvertSwitchCase = _ConvertSwitchCase + "default: ";
        }
        else
        {
            RecordLeft(ref L, "Case ");
            if (Right(L, 1) == ":") L = Left(L, Len(L) - 1);
            foreach (var iterV in new List<string>(Split(L, ", ")))
            {
                V = iterV;
                V = Trim(V);
                if (InStr(V, " To ") > 0)
                {
                    _ConvertSwitchCase = _ConvertSwitchCase + "default: /* TODO: Cannot Convert Ranged Case: " + L + " */";
                }
                else if (StartsWith(V, "Is "))
                {
                    _ConvertSwitchCase = _ConvertSwitchCase + "default: /* TODO: Cannot Convert Expression Case: " + L + " */";
                }
                else
                {
                    _ConvertSwitchCase = _ConvertSwitchCase + "case " + ConvertExpression(V) + ": ";
                }
            }
        }
        return _ConvertSwitchCase;
    }
    public static string ConvertWhile(string L)
    {
        string _ConvertWhile = "";
        string Exp = "";
        bool Closing = false;
        bool Invert = false;
        L = LTrim(L);
        if (RecordLeft(ref L, "Do While "))
        {
            Exp = L;
        }
        else if (RecordLeft(ref L, "Do Until "))
        {
            Exp = L;
            Invert = true;
        }
        else if (RecordLeft(ref L, "While "))
        {
            Exp = L;
        }
        else if (RecordLeft(ref L, "Loop While "))
        {
            Exp = L;
            Closing = true;
        }
        else if (RecordLeft(ref L, "Loop Until "))
        {
            Exp = L;
            Closing = true;
            Invert = true;
        }
        _ConvertWhile = "";
        if (Closing) _ConvertWhile = _ConvertWhile + "} ";
        _ConvertWhile = _ConvertWhile + "while(";
        if (Invert) _ConvertWhile = _ConvertWhile + "!(";
        _ConvertWhile = _ConvertWhile + ConvertExpression(Exp);
        if (Invert) _ConvertWhile = _ConvertWhile + ")";
        _ConvertWhile = _ConvertWhile + ")";
        if (!Closing) _ConvertWhile = _ConvertWhile + " {"; _ConvertWhile = _ConvertWhile + ";";
        return _ConvertWhile;
    }
    public static string ConvertFor(string L)
    {
        string _ConvertFor = "";
        string Var = "";
        string ForFrom = "";
        string ForTo = "";
        string ForStep = "";
        bool ForReverse = false;
        string ForCheck = "";
        L = Trim(L);
        RecordLeft(ref L, "For ");
        Var = RemoveUntil(ref L, " = ", true);
        ForFrom = RemoveUntil(ref L, " To ", true);
        ForTo = L;
        ForStep = RemoveUntil(ref ForTo, " Step ", true);
        if (ForStep == "") ForStep = "1";
        ForStep = ConvertExpression(ForStep);
        ForReverse = InStr(ForStep, "-") > 0;
        if (ForReverse) ForCheck = " >= "; ForCheck = " <= ";
        _ConvertFor = "";
        _ConvertFor = _ConvertFor + "for (";
        _ConvertFor = _ConvertFor + ExpandToken(Var) + " = " + ConvertExpression(ForFrom) + "; ";
        _ConvertFor = _ConvertFor + ExpandToken(Var) + ForCheck + ConvertExpression(ForTo) + "; ";
        _ConvertFor = _ConvertFor + ExpandToken(Var) + " += " + ForStep;
        _ConvertFor = _ConvertFor + ") {";
        return _ConvertFor;
    }
    public static string ConvertForEach(string L)
    {
        string _ConvertForEach = "";
        string Var = "";
        string ForSource = "";
        L = Trim(L);
        RecordLeft(ref L, "For ");
        RecordLeft(ref L, "Each ");
        Var = RemoveUntil(ref L, " In ", true);
        ForSource = L;
        _ConvertForEach = _ConvertForEach + "foreach (var iter" + Var + " in " + ConvertExpression(ForSource) + ") {" + vbCrLf + Var + " = iter" + Var + ";";
        return _ConvertForEach;
    }
    public static string ConvertType(string L)
    {
        string _ConvertType = "";
        bool isPrivate = false;
        bool isPublic = false;
        isPublic = RecordLeft(ref L, "Public ");
        isPrivate = RecordLeft(ref L, "Private ");
        RecordLeft(ref L, "Type ");
        CurrentTypeName = L;
        _ConvertType = "";
        if (!isPrivate) _ConvertType = _ConvertType + "public ";
        _ConvertType = _ConvertType + "class "; // `struct ` is available, but leads to non-conforming behavior when indexing in lists...
        _ConvertType = _ConvertType + L;
        _ConvertType = _ConvertType + "{ ";
        return _ConvertType;
    }
    public static string ConvertTypeLine(string L, CodeType vCodeType)
    {
        string _ConvertTypeLine = "";
        _ConvertTypeLine = ConvertDeclaration(L, DeclarationType.DECL_TYPE, vCodeType);
        return _ConvertTypeLine;
    }
    public static string ConvertEnum(string L)
    {
        string _ConvertEnum = "";
        bool isPrivate = false;
        bool isPublic = false;
        isPublic = RecordLeft(ref L, "Public ");
        isPrivate = RecordLeft(ref L, "Private ");
        RecordLeft(ref L, "Enum ");
        CurrentEnumName = L;
        _ConvertEnum = "";
        if (!isPrivate) _ConvertEnum = _ConvertEnum + "public ";
        _ConvertEnum = _ConvertEnum + "enum ";
        _ConvertEnum = _ConvertEnum + L;
        _ConvertEnum = _ConvertEnum + "{ ";
        return _ConvertEnum;
    }
    public static string ConvertEnumLine(string L)
    {
        string _ConvertEnumLine = "";
        string Name = "";
        string Value = "";
        List<string> Parts = new List<string>();
        Parts = new List<string>(Split(L, " = "));
        Name = Trim(Parts[0]);
        if (Parts.Count >= 1) Value = Trim(Parts[1]); Value = "";
        _ConvertEnumLine = "";
        if (Right(CurrentEnumName, 1) == "+") _ConvertEnumLine = _ConvertEnumLine + ", ";
        _ConvertEnumLine = _ConvertEnumLine + Name;
        if (Value != "") _ConvertEnumLine = _ConvertEnumLine + " = " + ConvertExpression(Value);
        CurrentEnumName = CurrentEnumName + "+"; // convenience
        return _ConvertEnumLine;
    }
    public static string ConvertProperty(string L, string FullContents, CodeType vCodeType)
    {
        string _ConvertProperty = "";
        string Name = "";
        int IX = 0;
        bool isPrivate = false;
        string ReturnType = "";
        string Discard = "";
        string PropertyType = "";
        string GetContents = "";
        string SetContents = "";
        IX = InStr(L, "(");
        Name = Left(L, IX - 1);
        RecordLeft(ref L, "Public ");
        isPrivate = RecordLeft(ref L, "Private ");
        RecordLeft(ref L, "Property ");
        RecordLeft(ref L, "Get ");
        RecordLeft(ref L, "Let ");
        RecordLeft(ref L, "Set ");
        IX = InStr(L, "(");
        Name = Left(L, IX - 1);
        if (InStr(ModuleProperties, Name) > 0) return _ConvertProperty;
        CurrentFunctionName = Name;
        CurrentFunctionReturnValue = "_" + Name;
        ModuleProperties = ModuleProperties + "[" + Name + "]";
        GetContents = FindPropertyBody(FullContents, "Get", Name, ref ReturnType);
        if (GetContents != "") GetContents = ConvertContents(GetContents, vCodeType, true);
        if (ReturnType == "") ReturnType = "Variant";
        SetContents = FindPropertyBody(FullContents, "Let", Name, ref Discard);
        if (SetContents == "") SetContents = FindPropertyBody(FullContents, "Set", Name, ref Discard);
        if (SetContents != "") SetContents = ConvertContents(SetContents, vCodeType, true);
        PropertyType = ConvertArgType(Name, ReturnType);
        _ConvertProperty = "";
        _ConvertProperty = _ConvertProperty + IIf(isPrivate, "private ", "public ");
        _ConvertProperty = _ConvertProperty + IIf(vCodeType == CodeType.CODE_MODULE, "static ", "");
        _ConvertProperty = _ConvertProperty + PropertyType + " " + Name + "{ " + vbCrLf;
        if (GetContents != "")
        {
            _ConvertProperty = _ConvertProperty + "get {" + vbCrLf;
            _ConvertProperty = _ConvertProperty + PropertyType + " " + CurrentFunctionReturnValue + ";" + vbCrLf;
            _ConvertProperty = _ConvertProperty + GetContents;
            _ConvertProperty = _ConvertProperty + "return " + CurrentFunctionReturnValue + ";" + vbCrLf;
            _ConvertProperty = _ConvertProperty + "}" + vbCrLf;
        }
        if (SetContents != "")
        {
            _ConvertProperty = _ConvertProperty + "set {" + vbCrLf;
            _ConvertProperty = _ConvertProperty + SetContents;
            _ConvertProperty = _ConvertProperty + "}" + vbCrLf;
        }
        _ConvertProperty = _ConvertProperty + "}" + vbCrLf;
        return _ConvertProperty;
    }
    public static string FindPropertyBody(string FullContents, string Typ, string Name, ref string ReturnType)
    {
        string _FindPropertyBody = "";
        int X = 0;
        X = InStr(FullContents, "Property " + Typ + " " + Name);
        if (X == 0) return _FindPropertyBody;
        _FindPropertyBody = Mid(FullContents, X);
        X = RegExNPos(_FindPropertyBody, "\\bEnd Property\\b", 0);
        _FindPropertyBody = Trim(Left(_FindPropertyBody, X - 1));
        RecordLeft(ref _FindPropertyBody, "Property " + Typ + " " + Name);
        RecordLeft(ref _FindPropertyBody, "(");
        X = 1;
        while (X > 0)
        {
            if (Left(_FindPropertyBody, 1) == "(") X = X + 1;
            if (Left(_FindPropertyBody, 1) == ")") X = X - 1;
            _FindPropertyBody = Mid(_FindPropertyBody, 2);
        }
        _FindPropertyBody = Trim(_FindPropertyBody);
        if (StartsWith(_FindPropertyBody, "As "))
        {
            _FindPropertyBody = Mid(_FindPropertyBody, 4);
            X = SearchLeft(1, _FindPropertyBody, ": " + vbCrLf, false, false);
            ReturnType = Left(_FindPropertyBody, X - 1);
            _FindPropertyBody = Mid(_FindPropertyBody, X);
        }
        while (StartsWith(_FindPropertyBody, vbCrLf)) { _FindPropertyBody = Mid(_FindPropertyBody, 3); }
        while (Right(_FindPropertyBody, 2) == vbCrLf) { _FindPropertyBody = Left(_FindPropertyBody, Len(_FindPropertyBody) - 2); }
        return _FindPropertyBody;
    }
    public static string ConvertDeclaration(string L, DeclarationType declType, CodeType vCodeType)
    {
        string _ConvertDeclaration = "";
        bool IsDim = false;
        bool isPrivate = false;
        bool isPublic = false;
        bool IsConst = false;
        bool isGlobal = false;
        bool isStatic = false;
        bool IsOptional = false;
        bool IsByVal = false;
        bool IsByRef = false;
        bool IsParamArray = false;
        bool IsWithEvents = false;
        bool IsEvent = false;
        int FixedLength = 0;
        bool IsNewable = false;
        L = Trim(L);
        if (L == "") return _ConvertDeclaration;
        IsDim = RecordLeft(ref L, "Dim ");
        isPrivate = RecordLeft(ref L, "Private ");
        isPublic = RecordLeft(ref L, "Public ");
        isGlobal = RecordLeft(ref L, "Global ");
        IsConst = RecordLeft(ref L, "Const ");
        isStatic = RecordLeft(ref L, "Static ");
        // If IsInStr(L, __S1) Then Stop
        if (isStatic && declType == DeclarationType.DECL_LOCAL) LineComment = LineComment + " TODO: (NOT SUPPORTED) C# Does not support static local variables.";
        dynamic Item = null;
        string LL = "";
        foreach (var iterItem in new List<string>(Split(L, ", ")))
        {
            Item = iterItem;
            int IX = 0;
            string ArgName = "";
            string ArgType = "";
            string ArgDefault = "";
            bool IsArray = false;
            string ArgTargetType = "";
            bool StandardEvent = false;
            if (_ConvertDeclaration != "" && declType != DeclarationType.DECL_SIGNATURE && declType != DeclarationType.DECL_EXTERN) _ConvertDeclaration = _ConvertDeclaration + vbCrLf;
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
            if (StartsWith(ArgType, "New "))
            {
                IsNewable = true;
                RecordLeft(ref ArgType, "New ");
                LineComment = LineComment + "TODO: (NOT SUPPORTED) Dimmable 'New' not supported on variable declaration.  Instantiated only on declaration.  Please ensure usages";
            }
            if (InStr(ArgType, " * ") > 0)
            {
                FixedLength = ValI(Trim(Mid(ArgType, InStr(ArgType, " * ") + 3)));
                ArgType = RemoveUntil(ref ArgType, " * ");
                LineComment = LineComment + "TODO: (NOT SUPPORTED) Fixed Length String not supported: " + ArgName + "(" + FixedLength + ")";
            }
            ArgName = LL;
            if (Right(ArgName, 2) == "()")
            {
                IsArray = true;
                ArgName = Left(ArgName, Len(ArgName) - 2);
            }
            else if (RegExTest(ArgName, "^[a-zA-Z_][a-zA-Z_0-9]*\\(.* To .*\\)$"))
            {
                IsArray = true;
                LineComment = LineComment + " TODO: (NOT SUPPORTED) Array ranges not supported: " + ArgName;
                ArgName = RemoveUntil(ref ArgName, "(");
            }
            else
            {
                IsArray = false;
            }
            ArgTargetType = ConvertArgType(ArgName, ArgType);
            StandardEvent = IsStandardEvent(ArgName, ArgType);
            switch (((declType)))
            {
                case DeclarationType.DECL_GLOBAL:  // global
                    if (isPublic || IsDim)
                    {
                        _ConvertDeclaration = _ConvertDeclaration + "public ";
                        if (vCodeType == CodeType.CODE_MODULE && !IsConst) _ConvertDeclaration = _ConvertDeclaration + "static ";
                    }
                    else
                    {
                        _ConvertDeclaration = _ConvertDeclaration + "public " + IIf(!IsConst, "static ", "");
                    }
                    if (IsConst) _ConvertDeclaration = _ConvertDeclaration + "const ";
                    _ConvertDeclaration = _ConvertDeclaration + IIf(IsArray, "List<" + ArgTargetType + ">", ArgTargetType) + " ";
                    _ConvertDeclaration = _ConvertDeclaration + ArgName;
                    if (ArgDefault != "")
                    {
                        _ConvertDeclaration = _ConvertDeclaration + " = " + ConvertExpression(ArgDefault);
                    }
                    else
                    {
                        _ConvertDeclaration = _ConvertDeclaration + " = " + ArgTypeDefault(ArgTargetType, IsArray, IsNewable); // VB6 always initializes variables on declaration
                    }
                    _ConvertDeclaration = _ConvertDeclaration + ";";
                    if (IsArray) ModuleArrays = ModuleArrays + "[" + ArgName + "]";
                    break;
                case DeclarationType.DECL_LOCAL:  // function contents
                    _ConvertDeclaration = _ConvertDeclaration + IIf(IsArray, "List<" + ArgTargetType + ">", ArgTargetType) + " ";
                    _ConvertDeclaration = _ConvertDeclaration + ArgName;
                    if (ArgDefault != "")
                    {
                        _ConvertDeclaration = _ConvertDeclaration + " = " + ConvertExpression(ArgDefault);
                    }
                    else
                    {
                        _ConvertDeclaration = _ConvertDeclaration + " = " + ArgTypeDefault(ArgTargetType, IsArray, IsNewable); // VB6 always initializes variables on declaration
                    }
                    _ConvertDeclaration = _ConvertDeclaration + ";";
                    if (IsArray) CurrentFunctionArrays = CurrentFunctionArrays + "[" + ArgName + "]";
                    CurrentFunctionArgs = CurrentFunctionArgs + "[" + ArgName + "]";
                    break;
                case DeclarationType.DECL_SIGNATURE:  // sig def
                    if (_ConvertDeclaration != "") _ConvertDeclaration = _ConvertDeclaration + ", ";
                    if (IsByRef || !IsByVal) _ConvertDeclaration = _ConvertDeclaration + "ref ";
                    _ConvertDeclaration = _ConvertDeclaration + IIf(IsArray, "List<" + ArgTargetType + ">", ArgTargetType) + " ";
                    _ConvertDeclaration = _ConvertDeclaration + ArgName;
                    if (ArgDefault != "") _ConvertDeclaration = _ConvertDeclaration + " = " + ConvertExpression(ArgDefault); // default on method sig means optional param
                    if (IsArray) CurrentFunctionArrays = CurrentFunctionArrays + "[" + ArgName + "]";
                    CurrentFunctionArgs = CurrentFunctionArgs + "[" + ArgName + "]";
                    break;
                case DeclarationType.DECL_TYPE:
                    _ConvertDeclaration = _ConvertDeclaration + "public " + ArgTargetType + " " + ArgName + ";";
                    break;
                case DeclarationType.DECL_ENUM:
                    break;
                case DeclarationType.DECL_EXTERN:
                    if (_ConvertDeclaration != "") _ConvertDeclaration = _ConvertDeclaration + ", ";
                    if (IsByRef || !IsByVal) _ConvertDeclaration = _ConvertDeclaration + "ref ";
                    _ConvertDeclaration = _ConvertDeclaration + IIf(IsArray, "List<" + ArgTargetType + ">", ArgTargetType) + " ";
                    _ConvertDeclaration = _ConvertDeclaration + ArgName;
                    break;
            }
            // If IsParamArray Then Stop
            if (ArgType == "" && !IsEvent && !StandardEvent)
            {
            }
            if (declType == DeclarationType.DECL_SIGNATURE)
            {
                if (IsParamArray)
                {
                }
                else
                {
                    if (!IsByVal && !IsByRef && !StandardEvent)
                    {
                    }
                }
                if (IsOptional && IsByRef)
                {
                }
                if (IsOptional && ArgDefault == "")
                {
                }
            }
        }
        return _ConvertDeclaration;
    }
    // Function IsStandardEvent(ByVal ArgName As String, ByVal ArgType As String) As Boolean
    // If ArgName = __S1 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // If ArgName = __S1 And ArgType = __S2 Then IsStandardEvent = True: Exit Function
    // IsStandardEvent = False
    // End Function
    public static string ConvertArgType(string Name, string Typ)
    {
        string _ConvertArgType = "";
        switch (Typ)
        {
            case "Long":
            case "Integer":
            case "Int32":
            case "Short":
                _ConvertArgType = "int";
                break;
            case "Currency":
                _ConvertArgType = "decimal";
                break;
            case "Date":
                _ConvertArgType = "DateTime";
                break;
            case "Double":
            case "Float":
            case "Single":
                _ConvertArgType = "double";
                break;
            case "String":
                _ConvertArgType = "string";
                break;
            case "Boolean":
                _ConvertArgType = "bool";
                break;
            case "Variant":
            case "Object":
                _ConvertArgType = "dynamic";
                break;
            default:
                _ConvertArgType = Typ;
                break;
        }
        return _ConvertArgType;
    }
    public static string ArgTypeDefault(string ArgType, bool AsArray = false, bool IsNewable = false)
    {
        string _ArgTypeDefault = "";
        if (!AsArray)
        {
            switch (LCase(ArgType))
            {
                case "string":
                    _ArgTypeDefault = "\"\"";
                    break;
                case "long":
                case "int":
                case "integer":
                case "short":
                case "byte":
                case "decimal":
                case "float":
                case "double":
                case "currency":
                    _ArgTypeDefault = "0";
                    break;
                case "boolean":
                case "bool":
                    _ArgTypeDefault = "false";
                    break;
                case "vbtristate":
                    _ArgTypeDefault = "vbUseDefault";
                    break;
                case "datetime":
                case "date":
                    _ArgTypeDefault = "DateTime.MinValue";
                    break;
                default:
                    _ArgTypeDefault = (IsNewable ? "new " + ArgType + "()" : "null");
                    break;
            }
        }
        else
        {
            _ArgTypeDefault = "new List<" + ArgType + ">()";
        }
        return _ArgTypeDefault;
    }
    public static string ConvertSignature(string LL, CodeType vCodeType = CodeType.CODE_FORM)
    {
        string _ConvertSignature = "";
        string L = "";
        bool WithReturn = false;
        bool isPublic = false;
        bool isPrivate = false;
        bool IsFriend = false;
        bool IsPropertyGet = false;
        bool IsPropertyLet = false;
        bool IsPropertySet = false;
        bool IsFunction = false;
        bool IsSub = false;
        L = LL;
        isPrivate = RecordLeft(ref L, "Private ");
        isPublic = RecordLeft(ref L, "Public ");
        IsFriend = RecordLeft(ref L, "Friend ");
        IsSub = RecordLeft(ref L, "Sub ");
        IsFunction = RecordLeft(ref L, "Function ");
        IsPropertyGet = RecordLeft(ref L, "Property Get ");
        IsPropertyLet = RecordLeft(ref L, "Property let ");
        IsPropertySet = RecordLeft(ref L, "Property set ");
        WithReturn = IsFunction || IsPropertyGet;
        int IX = 0;
        int Ix2 = 0;
        string Name = "";
        string Args = "";
        string Ret = "";
        string RetTargetType = "";
        bool IsArray = false;
        IX = InStr(L, "(");
        if (IX == 0) return _ConvertSignature;
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
        Ret = Replace(Ret, " As ", "");
        IsArray = Right(Ret, 2) == "()";
        if (IsArray) Ret = Left(Ret, Len(Ret) - 2);
        RetTargetType = ConvertArgType(Name, Ret);
        if (IsArray) RetTargetType = "List<" + RetTargetType + ">";
        CurrentFunctionName = Name;
        CurrentFunctionReturnValue = (WithReturn ? "_" + CurrentFunctionName : "");
        _ConvertSignature = "";
        if (isPublic) _ConvertSignature = _ConvertSignature + "public ";
        if (isPrivate) _ConvertSignature = _ConvertSignature + "private ";
        if (vCodeType == CodeType.CODE_MODULE) _ConvertSignature = _ConvertSignature + "static ";
        _ConvertSignature = _ConvertSignature + IIf(Ret == "", "void ", RetTargetType + " ");
        _ConvertSignature = _ConvertSignature + Name + "(" + ConvertDeclaration(Args, DeclarationType.DECL_SIGNATURE, vCodeType) + ") {";
        if (WithReturn)
        {
            _ConvertSignature = _ConvertSignature + vbCrLf + RetTargetType + " " + CurrentFunctionReturnValue + " = " + ArgTypeDefault(RetTargetType) + ";";
        }
        return _ConvertSignature;
    }
    public static string ConvertDeclare(string L)
    {
        string _ConvertDeclare = "";
        // Private Declare Function CreateFile Lib __S1 Alias __S2 (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
        // [DllImport(__S1)]
        // public static extern int MessageBox(int h, string m, string c, int type);
        bool isPrivate = false;
        bool isPublic = false;
        bool IsFunction = false;
        bool IsSub = false;
        int X = 0;
        string Name = "";
        string cLib = "";
        string cAlias = "";
        string Args = "";
        string Ret = "";
        L = Trim(L);
        isPrivate = RecordLeft(ref L, "Private ");
        isPublic = RecordLeft(ref L, "Public ");
        L = StripLeft(L, "Declare ");
        IsFunction = RecordLeft(ref L, "Function ");
        IsSub = RecordLeft(ref L, "Sub ");
        X = InStr(L, " ");
        Name = Left(L, X - 1);
        L = Mid(L, X + 1);
        if (RecordLeft(ref L, "Lib "))
        {
            X = InStr(L, " ");
            cLib = Left(L, X - 1);
            // If Left(cLib, 1) = __S1 Then cLib = Mid(cLib, 2, Len(cLib) - 2)
            L = Mid(L, X + 1);
        }
        if (RecordLeft(ref L, "Alias "))
        {
            X = InStr(L, " ");
            cAlias = Left(L, X - 1);
            // If Left(cAlias, 1) = __S1 Then cAlias = Mid(cAlias, 2, Len(cAlias) - 2)
            L = Mid(L, X + 1);
        }
        X = InStrRev(L, ")");
        Ret = Trim(Mid(L, X + 1));
        Ret = Replace(Ret, "As ", "");
        Args = Mid(L, 2, X - 2);
        _ConvertDeclare = "";
        _ConvertDeclare = _ConvertDeclare + "[DllImport(" + cLib + ")]" + vbCrLf;
        _ConvertDeclare = _ConvertDeclare + IIf(isPrivate, "private ", "public ") + "static extern ";
        _ConvertDeclare = _ConvertDeclare + IIf(Ret == "", "void", ConvertArgType("return", Ret)) + " ";
        _ConvertDeclare = _ConvertDeclare + Name + "(";
        _ConvertDeclare = _ConvertDeclare + ConvertDeclaration(Args, DeclarationType.DECL_EXTERN, CodeType.CODE_MODULE);
        _ConvertDeclare = _ConvertDeclare + ");";
        return _ConvertDeclare;
    }
    public static List<string> SplitByComma(string L)
    {
        List<string> _SplitByComma = null;
        List<string> Results = new List<string>();
        int ResultCount = 0;
        int N = 0;
        int I = 0;
        string C = "";
        int Depth = 0;
        string Part = "";
        N = Len(L);
        for (I = 1; I <= N; I += 1)
        {
            C = Mid(L, I, 1);
            if (C == "(")
            {
                Depth = Depth + 1;
                Part = Part + C;
            }
            else if (Depth > 0 && C == ")")
            {
                Depth = Depth - 1;
                Part = Part + C;
            }
            else if (Depth == 0 && (C == "," || C == ")"))
            {
                ResultCount = ResultCount + 1;
                // TODO: (NOT SUPPORTED): ReDim Preserve Results(1 To ResultCount)
                Results[ResultCount] = Trim(Part);
                Part = "";
            }
            else
            {
                Part = Part + C;
            }
        }
        ResultCount = ResultCount + 1;
        // TODO: (NOT SUPPORTED): ReDim Preserve Results(1 To ResultCount)
        Results[ResultCount] = Trim(Part);
        _SplitByComma = Results;
        return _SplitByComma;
    }
    public static int FindNextOperator(string L)
    {
        int _FindNextOperator = 0;
        int N = 0;
        N = Len(L);
        for (_FindNextOperator = 1; _FindNextOperator <= N; _FindNextOperator += 1)
        {
            if (StartsWith(Mid(L, _FindNextOperator), " && ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " || ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " ^^ ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " - ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " + ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " * ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " / ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " < ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " > ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " >= ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " <= ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " != ")) return _FindNextOperator;
            if (StartsWith(Mid(L, _FindNextOperator), " == ")) return _FindNextOperator;
        }
        _FindNextOperator = 0;
        return _FindNextOperator;
    }
    public static string ConvertIIf(string L)
    {
        string _ConvertIIf = "";
        List<string> Parts = new List<string>();
        string Condition = "";
        string TruePart = "";
        string FalsePart = "";
        Parts = SplitByComma(Mid(Trim(L), 5, Len(L) - 5));
        Condition = Parts[1];
        TruePart = Parts[2];
        FalsePart = Parts[3];
        _ConvertIIf = "(" + ConvertExpression(Condition) + " ? " + ConvertExpression(TruePart) + " : " + ConvertExpression(FalsePart) + ")";
        return _ConvertIIf;
    }
    public static string ConvertStatement(string L)
    {
        string _ConvertStatement = "";
        bool NonCodeLine = false;
        L = Trim(L);
        if (StartsWith(L, "Set ")) L = Mid(L, 5);
        if (StartsWith(L, "Option "))
        {
            // ignore __S1 directives
            NonCodeLine = true;
        }
        else if (RegExTest(L, "^[ ]*Exit (Function|Sub|Property)$"))
        {
            _ConvertStatement = _ConvertStatement + "return";
            if (CurrentFunctionReturnValue != "") _ConvertStatement = _ConvertStatement + " " + CurrentFunctionReturnValue;
        }
        else if (RegExTest(L, "^[ ]*Exit (Do|Loop|For|While)$"))
        {
            _ConvertStatement = _ConvertStatement + "break";
        }
        else if (InStr(L, " = ") > 0)
        {
            int IX = 0;
            string AssignmentTarget = "";
            string AssignmentValue = "";
            IX = InStr(L, " = ");
            AssignmentTarget = Trim(Left(L, IX - 1));
            if (InStr(AssignmentTarget, "(") > 0) AssignmentTarget = ConvertExpression(AssignmentTarget);
            if (IsControlRef(AssignmentTarget, ModuleName()))
            {
                // If InStr(AssignmentTarget, __S1) > 0 Then Stop
                AssignmentTarget = modRefScan.FormControlRepl(AssignmentTarget, ModuleName());
            }
            if (AssignmentTarget == CurrentFunctionName) AssignmentTarget = CurrentFunctionReturnValue;
            AssignmentValue = Mid(L, IX + 3);
            _ConvertStatement = AssignmentTarget + " = " + ConvertExpression(AssignmentValue);
        }
        else if (RegExTest(L, "^[ ]*Unload "))
        {
            L = Trim(L);
            RecordLeft(ref L, "Unload ");
            _ConvertStatement = (L == "Me" ? "Unload()" : L + ".instance.Unload()");
        }
        else if (RegExTest(L, "^[ ]*With") || RegExTest(L, "^[ ]*End With"))
        {
            _ConvertStatement = "// TODO: (NOT SUPPORTED): " + L;
            NonCodeLine = true;
        }
        else if (RegExTest(L, "^[ ]*(On Error|Resume) "))
        {
            _ConvertStatement = "// TODO: (NOT SUPPORTED): " + L;
            NonCodeLine = true;
        }
        else if (RegExTest(L, "^[ ]*ReDim "))
        {
            _ConvertStatement = "// TODO: (NOT SUPPORTED): " + L;
            NonCodeLine = true;
        }
        else if (RegExTest(L, "^[ ]*Err.Clear"))
        {
            _ConvertStatement = "// TODO: (NOT SUPPORTED): " + L;
            NonCodeLine = true;
        }
        else if (RegExTest(L, "^[ ]*(([a-zA-Z_()0-9.]\\.)*)?[a-zA-Z_0-9.]+$"))
        { // Method call without parens or args (statement, not expression)
            _ConvertStatement = _ConvertStatement + L + "()";
        }
        else if (RegExTest(L, "^[ ]*(([a-zA-Z_()0-9.]\\.)*)?[a-zA-Z_0-9.]+ .*"))
        { // Method call without parens but with args (statement, not expression)
            string FunctionCall = "";
            string ArgList = "";
            dynamic ArgPart = null;
            int ArgN = 0;
            FunctionCall = RegExNMatch(L, "^[ ]*((([a-zA-Z_()0-9.]\\.)*)?[a-zA-Z_0-9.]+)", 0);
            ArgList = Trim(Mid(L, Len(FunctionCall) + 1));
            _ConvertStatement = ExpandFunctionCall(FunctionCall, ArgList);
        }
        else
        {
            _ConvertStatement = L;
        }
        if (!NonCodeLine) _ConvertStatement = _ConvertStatement + ";";
        return _ConvertStatement;
    }
    public static string ConvertExpression(string L)
    {
        string _ConvertExpression = "";
        L = Replace(L, " \\ ", " / ");
        L = Replace(L, " = ", " == ");
        L = Replace(L, " Mod ", " % ");
        L = Replace(L, " & ", " + ");
        L = Replace(L, " And ", " && ");
        L = Replace(L, " Or ", " || ");
        L = Replace(L, " Xor ", " ^^ ");
        L = Replace(L, " Is ", " == ");
        if (InStr(L, " Like ") > 0) LineComment = LineComment + "TODO: (NOT SUPPORTED) LIKE statement changed to ==: " + L;
        L = Replace(L, " Like ", " == ");
        L = Replace(L, " <> ", " != ");
        L = RegExReplace(L, "\\bNot\\b", "!");
        L = RegExReplace(L, "\\bFalse\\b", "false");
        L = RegExReplace(L, "\\bTrue\\b", "true");
        if (LMatch(LTrim(L), "New ")) L = "new " + Mid(LTrim(L), 5) + "()";
        if (StartsWith(L, "IIf("))
        {
            L = ConvertIIf(L);
        }
        else
        {
            L = ParseAndExpandExpression(L);
        }
        if (CurrentFunctionName != "") L = RegExReplace(L, "\\b" + CurrentFunctionName + "([^(a-zA-Z_])", CurrentFunctionReturnValue + "$1");
        _ConvertExpression = L;
        return _ConvertExpression;
    }
    public static string ParseAndExpandExpression(string Src)
    {
        string _ParseAndExpandExpression = "";
        string S = "";
        string Token = "";
        string T = "";
        int I = 0;
        int J = 0;
        int X = 0;
        int Y = 0;
        string C = "";
        string FunctionName = "";
        string FunctionArgs = "";
        Token = EXPRESSION_TOKEN_PREFIX + CLng(Rnd() * 1000000);
        S = RegExNMatch(Src, "\\([^()]+\\)", 0);
        if (S != "")
        {
            X = InStr(Src, S);
            Src = Replace(Src, S, Token, 1, 1);
            if (X > 1) C = Mid(Src, X - 1, 1); C = "";
            if (X > 1 && C != "(" && C != ")" && C != " ")
            {
                Y = SearchLeft(X - 1, Src, "() ", false, true);
                FunctionName = Mid(Src, Y + 1, X - Y - 1);
                Src = Replace(Src, FunctionName + Token, Token, 1, 1);
                FunctionArgs = Mid(S, 2, Len(S) - 2);
                if (modRefScan.IsControlRef(FunctionName, ModuleName()))
                {
                    _ParseAndExpandExpression = FunctionName + "[" + FunctionArgs + "]" + "." + ConvertControlProperty("", "", FormControlRefDeclType(FunctionName, ModuleName()));
                    return _ParseAndExpandExpression;
                }
                FunctionName = ExpandToken(FunctionName, true);
                S = ExpandFunctionCall(FunctionName, FunctionArgs);
                _ParseAndExpandExpression =ParseAndExpandExpression(Src);
                _ParseAndExpandExpression = Replace(_ParseAndExpandExpression, Token, S);
                // Debug.Print __S1 & S
                return _ParseAndExpandExpression;
            }
            else
            { // not a function, but sub expression maybe math
                T = Mid(S, 2, Len(S) - 2);
                X = FindNextOperator(T);
                if (X == 0)
                {
                    _ParseAndExpandExpression = ExpandToken(T);
                }
                else
                {
                    Y = InStr(X + 2, T, " ");
                    S = ExpandToken(Left(T, X - 1)) + Mid(T, X, Y - X + 1) + ParseAndExpandExpression(Mid(T, Y + 1));
                }
                _ParseAndExpandExpression = ParseAndExpandExpression(Src);
                _ParseAndExpandExpression = Replace(_ParseAndExpandExpression, Token, "(" + S + ")");
                // Debug.Print __S1 & S
                return _ParseAndExpandExpression;
            }
        }
        // no subexpression.  Check for math
        X = FindNextOperator(Src);
        if (X == 0)
        {
            _ParseAndExpandExpression = ExpandToken(Src);
            // Debug.Print __S1 & S
            return _ParseAndExpandExpression;
        }
        else
        {
            Y = InStr(X + 2, Src, " ");
            _ParseAndExpandExpression = ParseAndExpandExpression(Left(Src, X - 1)) + Mid(Src, X, Y - X + 1) + ParseAndExpandExpression(Mid(Src, Y + 1));
            // Debug.Print __S1 & S
            return _ParseAndExpandExpression;
        }
        return _ParseAndExpandExpression;
    }
    public static string ExpandToken(string T, bool WillAddParens = false, bool AsLast = false)
    {
        string _ExpandToken = "";
        bool WithNot = false;
        WithNot = Left(T, 1) == "!";
        if (WithNot) T = Mid(T, 2);
        // If InStr(T, __S1) > 0 Then Stop
        // If InStr(T, __S1) > 0 Then Stop
        // If InStr(T, __S1) > 0 Then Stop
        // Debug.Print __S1 & T
        if (T == CurrentFunctionName)
        {
            T = CurrentFunctionReturnValue;
        }
        else if (T == "Rnd")
        {
            T = T + "()";
        }
        else if (T == "Me")
        {
            T = "this";
        }
        else if (T == "App.Path")
        {
            T = "AppContext.BaseDirectory";
        }
        else if (T == "Now")
        {
            T = "DateTime.Now";
        }
        else if (T == "Nothing")
        {
            T = "null";
        }
        else if (T == "Err.Number")
        {
            T = "Err().Number";
        }
        else if (T == "Err.Description")
        {
            T = "Err().Description";
        }
        else if (InStr(CurrentFunctionArgs, T) == 0 && !WillAddParens && (IsFuncRef(T) || IsLocalFuncRef(T)))
        {
            // Debug.Print __S1 & T
            T = T + "()";
        }
        else if (modRefScan.IsFormRef(T))
        {
            T = FormRefRepl(T);
        }
        else if (modRefScan.IsControlRef(T, ModuleName()))
        {
            T = FormControlRepl(T, ModuleName());
        }
        else if (modRefScan.IsEnumRef(T))
        {
            T = modRefScan.EnumRefRepl(T);
        }
        else if (Left(T, 2) == "&H")
        {
            T = "0x" + Mid(T, 3);
            if (Right(T, 1) == "&") T = Left(T, Len(T) - 1);
        }
        else if (RegExTest(T, "^[0-9.-]+&$"))
        {
            T = Left(T, Len(T) - 1);
        }
        else if (IsInStr(T, "."))
        {
            List<string> Parts = new List<string>();
            int I = 0;
            string Part = "";
            bool IsLast = false;
            string TOut = "";
            // Debug.Print __S1 & T
            TOut = "";
            Parts = new List<string>(Split(T, "."));
            for (I = 0; I <= Parts.Count; I += 1)
            {
                Part = Parts[I];
                IsLast = I == Parts.Count;
                if (TOut != "") TOut = TOut + ".";
                TOut = TOut + ExpandToken(Part, WillAddParens, IsLast);
            }
            T = TOut;
        }
        _ExpandToken = (WithNot ? "!" : "");
        return _ExpandToken;
    }
    public static string ExpandFunctionCall(string FunctionName, string Args)
    {
        string _ExpandFunctionCall = "";
        if (InStr(ModuleArrays + CurrentFunctionArrays + FormControlArrays, "[" + FunctionName + "]") > 0)
        {
            _ExpandFunctionCall = FunctionName + "[" + ProcessFunctionArgs(Args) + "]";
        }
        else if (FunctionName == "LBound")
        {
            _ExpandFunctionCall = "0";
        }
        else if (FunctionName == "UBound")
        {
            _ExpandFunctionCall = Args + ".Count";
        }
        else if (FunctionName == "Split")
        {
            _ExpandFunctionCall = "new List<string>(" + FunctionName + "(" + ProcessFunctionArgs(Args) + ")" + ")";
        }
        else if (FunctionName == "Debug.Print")
        {
            _ExpandFunctionCall = "Console.WriteLine(" + ProcessFunctionArgs(Args) + ")";
        }
        else if (FunctionName == "Erase")
        {
            _ExpandFunctionCall = Args + ".Clear()";
        }
        else if (FunctionName == "GoTo")
        {
            _ExpandFunctionCall = "goto " + Args;
        }
        else if (FunctionName == "Array")
        {
            _ExpandFunctionCall = "new List<dynamic>() {" + ProcessFunctionArgs(Args) + "}";
        }
        else if (FunctionName == "Show")
        {
            _ExpandFunctionCall = (Args == "" ? "Show()" : "ShowDialog()");
        }
        else if (modRefScan.IsFormRef(FunctionName))
        {
            _ExpandFunctionCall = modRefScan.FormRefRepl(FunctionName) + "(" + ProcessFunctionArgs(Args, FunctionName) + ")";
        }
        else
        {
            _ExpandFunctionCall = FunctionName + "(" + ProcessFunctionArgs(Args, FunctionName) + ")";
        }
        _ExpandFunctionCall = RegExReplace(_ExpandFunctionCall, "\\.Show\\(.+\\)", ".ShowDialog()");
        return _ExpandFunctionCall;
    }
    public static string ProcessFunctionArgs(string Args, string FunctionName = "")
    {
        string _ProcessFunctionArgs = "";
        dynamic Arg = null;
        int I = 0;
        foreach (var iterArg in SplitByComma(Args))
        {
            Arg = iterArg;
            I = I + 1;
            if (_ProcessFunctionArgs != "") _ProcessFunctionArgs = _ProcessFunctionArgs + ", ";
            if (FunctionName != "")
            {
                if (modRefScan.IsFuncRef(FunctionName))
                {
                    if (I <= FuncRefDeclArgCnt(FunctionName) && modRefScan.FuncRefArgByRef(FunctionName, I))
                    {
                        // If IsInStr(Arg, STRING_TOKEN_PREFIX) Then Stop
                        _ProcessFunctionArgs = _ProcessFunctionArgs + "ref ";
                    }
                }
            }
            cValP(null, "", "");
            _ProcessFunctionArgs = _ProcessFunctionArgs + ConvertExpression(Arg);
        }
        return _ProcessFunctionArgs;
    }

}
