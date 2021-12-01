using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using VB2CS.Forms;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static Microsoft.VisualBasic.VBMath;
using static modConfig;
using static modRegEx;
using static modTextFiles;
using static VBExtension;



static class modUtils
{
    public const string patToken = "([a-zA-Z_][a-zA-Z_0-9]*)";
    public const string patNotToken = "([^a-zA-Z_0-9])";
    public const string patTokenDot = "([a-zA-Z_.][a-zA-Z_0-9.]*)";
    public const string vbCrLf2 = vbCrLf + vbCrLf;
    public const string vbCrLf3 = vbCrLf + vbCrLf + vbCrLf;
    public const string vbCrLf4 = vbCrLf + vbCrLf + vbCrLf + vbCrLf;
    public const string STR_CHR_UCASE = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    public const string STR_CHR_LCASE = "abcdefghijklmnopqrstuvwxyz";
    public const string STR_CHR_DIGIT = "1234567890"; // eol comment
                                                      // block comment
                                                      // again
    public static bool IsInStr(string Src, string Find)
    {
        bool _IsInStr = false; _IsInStr = InStr(Src, Find) > 0; return _IsInStr;
    }
    // after
    public static bool IsNotInStr(string S, string Fnd)
    {
        bool _IsNotInStr = false; _IsNotInStr = !IsInStr(S, Fnd); return _IsNotInStr;
    }
    public static bool FileExists(string FN)
    {
        bool _FileExists = false; _FileExists = FN != "" && Dir(FN) != ""; return _FileExists;
    }
    public static bool DirExists(string FN)
    {
        bool _DirExists = false; _DirExists = FN != "" && Dir(FN, vbDirectory) != ""; return _DirExists;
    }
    public static string ProjFileName(string FN)
    {
        string _ProjFileName = ""; _ProjFileName = Mid(FN, InStrRev(FN, "\\") + 1); return _ProjFileName;
    }
    public static string FileBaseName(string FN)
    {
        string _FileBaseName = ""; _FileBaseName = Left(ProjFileName(FN), InStrRev(ProjFileName(FN), ".") - 1); return _FileBaseName;
    }
    public static string FilePath(string FN)
    {
        string _FilePath = ""; _FilePath = Left(FN, InStrRev(FN, "\\")); return _FilePath;
    }
    public static string ChgExt(string FN, string NewExt)
    {
        string _ChgExt = ""; _ChgExt = Left(FN, InStrRev(FN, ".") - 1) + NewExt; return _ChgExt;
    }
    public static string tLeft(string Str, int N)
    {
        string _tLeft = ""; _tLeft = Left(Trim(Str), N); return _tLeft;
    }
    public static string tMid(string Str, int N, int M = 0)
    {
        string _tMid = ""; _tMid = (M == 0 ? Mid(Trim(Str), N) : Mid(Trim(Str), N, M)); return _tMid;
    }
    public static int StrCnt(string Src, string Str)
    {
        int _StrCnt = 0; _StrCnt = (Len(Src) - Len(Replace(Src, Str, ""))) / IIf(Len(Str) == 0, 1, Len(Str)); return _StrCnt;
    }
    public static bool LMatch(string Src, string tMatch)
    {
        bool _LMatch = false; _LMatch = Left(Src, Len(tMatch)) == tMatch; return _LMatch;
    }
    public static bool tLMatch(string Src, string tMatch)
    {
        bool _tLMatch = false; _tLMatch = Left(LTrim(Src), Len(tMatch)) == tMatch; return _tLMatch;
    }
    public static int Px(int Twips)
    {
        int _Px = 0; _Px = Twips / 14; return _Px;
    }
    public static string Quote(string S)
    {
        string _Quote = ""; _Quote = "\"" + S + "\""; return _Quote;
    }
    public static string AlignString(string S, int N)
    {
        string _AlignString = ""; _AlignString = Left(S + Space(N), N); return _AlignString;
    }
    public static string Capitalize(string S)
    {
        string _Capitalize = ""; _Capitalize = UCase(Left(S, 1)) + Mid(S, 2); return _Capitalize;
    }
    public static string DevelopmentFolder()
    {
        string _DevelopmentFolder = ""; _DevelopmentFolder = AppContext.BaseDirectory + "\\"; return _DevelopmentFolder;
    }
    public static bool IsIDE()
    {
        bool _IsIDE = false;
        // IsIDE = False
        // Exit Function
        // works on a very simple princicple... debug statements don't get compiled...
        // TODO: (NOT SUPPORTED): On Error GoTo IDEInUse
        Console.WriteLine(1 / 0); // division by zero error
        _IsIDE = false;
        return _IsIDE;
    IDEInUse:;
        _IsIDE = true;
        return _IsIDE;
    }
    public static bool IsIn(string S, ref List<dynamic> K)
    {
        bool _IsIn = false;
        dynamic L = null;
        foreach (var iterL in K)
        {
            L = iterL;
            if (S == L) { _IsIn = true; return _IsIn; }
        }
        return _IsIn;
    }
    public static bool WriteOut(string F, string S, string O = "")
    {
        bool _WriteOut = false;
        if (!IsConverted(F, O))
        {
            _WriteOut = WriteFile(OutputFolder(O) + F, S, true);
        }
        else
        {
            Console.WriteLine("Already converted: " + F);
        }
        return _WriteOut;
    }
    public static bool IsConverted(string F, string O = "")
    {
        bool _IsConverted = false;
        _IsConverted = IsInStr(Left(ReadEntireFile(OutputFolder(O) + F), 100), "### CONVERTED");
        return _IsConverted;
    }
    public static string FileExt(string FN, bool vLCase = true)
    {
        string _FileExt = "";
        if (FN == "") return _FileExt;
        if (InStr(FN, ".") == 0) return _FileExt;
        _FileExt = Mid(FN, InStrRev(FN, "."));
        _FileExt = (vLCase ? LCase(_FileExt) : _FileExt);
        return _FileExt;
    }
    public static string deQuote(string Src)
    {
        string _deQuote = "";
        if (Left(Src, 1) == "\"") Src = Mid(Src, 2);
        if (Right(Src, 1) == "\"") Src = Left(Src, Len(Src) - 1);
        _deQuote = Src;
        return _deQuote;
    }
    public static string deWS(string S)
    {
        string _deWS = "";
        while (IsInStr(S, " " + vbCrLf))
        {
            S = Replace(S, " " + vbCrLf, vbCrLf);
        }
        while (IsInStr(S, vbCrLf4))
        {
            S = Replace(S, vbCrLf4, vbCrLf3);
        }
        S = Replace(S, "{" + vbCrLf2, "{" + vbCrLf);
        S = RegExReplace(S, "(" + vbCrLf2 + ")([ ]*{)", vbCrLf + "$2");
        S = RegExReplace(S, "([ ]*case .*:)" + vbCrLf2, "$1" + vbCrLf);
        _deWS = S;
        return _deWS;
    }
    public static string nlTrim(string Str)
    {
        string _nlTrim = "";
        while (InStr(" " + vbTab + vbCr + vbLf, Left(Str, 1)) != 0 && Str != "") { Str = Mid(Str, 2); }
        while (InStr(" " + vbTab + vbCr + vbLf, Right(Str, 1)) != 0 && Str != "") { Str = Mid(Str, 1, Len(Str) - 1); }
        _nlTrim = Str;
        return _nlTrim;
    }
    public static string sSpace(int N)
    {
        string _sSpace = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        _sSpace = Space(N);
        return _sSpace;
    }
    public static string nextBy(string Src, string Del = "\"", int Ind = 1, bool ProcessVBComments = false)
    {
        string _nextBy = "";
        int L = 0;
        DoEvents();
        L = InStr(Src, Del);
        if (L == 0) { _nextBy = (Ind <= 1 ? Src : ""); return _nextBy; }
        if (Ind <= 1)
        {
            _nextBy = Left(Src, L - 1);
        }
        else
        {
            _nextBy = _nextBy(Mid(Src, L + Len(Del)), Del, Ind - 1);
        }
        return _nextBy;
    }
    public static int StrQCnt(string Src, string Str)
    {
        int _StrQCnt = 0;
        int N = 0;
        int I = 0;
        string C = "";
        bool Q = false;
        _StrQCnt = 0;
        N = Len(Src);
        for (I = 1; I <= N; I += 1)
        {
            C = Mid(Src, I, 1);
            if (C == "\"")
            {
                Q = !Q;
            }
            else
            {
                if (!Q)
                {
                    if (LMatch(Mid(Src, I), Str)) _StrQCnt = _StrQCnt + 1;
                }
            }
        }
        return _StrQCnt;
    }
    public static int nextByPCt(string Src, string Del = "\"", int Ind = 1)
    {
        int _nextByPCt = 0;
        int M = 0;
        int N = 0;
        string F = "";
        N = 0;
        do
        {
            N = N + 1;
            if (N > 1000) break;
            F = nextByP(Src, Del, N);
            if (F == "")
            {
                M = M + 1;
                if (M >= 10) break;
            }
            else
            {
                M = 0;
            }
        } while (true);
        _nextByPCt = N - M;
        return _nextByPCt;
    }
    public static string nextByP(string Src, string Del = "\"", int Ind = 1)
    {
        string _nextByP = "";
        string F = "";
        int N = 0;
        int M = 0;
        string R = "";
        string T = "";
        N = 0;
        F = "";
        do
        {
            M = M + 1;
            if (M > 100) break;
            N = N + 1;
            T = nextBy(Src, Del, N);
            R = R + IIf(Len(R) == 0, "", Del) + T;
        } while (!(StrQCnt(R, "(") == StrQCnt(R, ")")));
        if (Ind <= 1)
        {
            _nextByP = R;
        }
        else
        {
            _nextByP = _nextByP(Mid(Src, Len(R) + Len(Del) + 1), Del, Ind - 1);
        }
        return _nextByP;
    }
    public static string NextByOp(string Src, int Ind = 1, ref string Op = "")
    {
        string _NextByOp = "";
        string A = "";
        string S = "";
        string D = "";
        string M = "";
        string C = "";
        string E = "";
        string I = "";
        string cNE = "";
        string cLT = "";
        string cGT = "";
        string cLE = "";
        string cGE = "";
        string cEQ = "";
        string lA = "";
        string lO = "";
        string lM = "";
        string LL = "";
        string xIs = "";
        string xLk = "";
        string P = "";
        int K = 0;
        A = nextByP(Src, " + ");
        S = nextByP(Src, " - ");
        M = nextByP(Src, " * ");
        D = nextByP(Src, " / ");
        I = nextByP(Src, " \\ ");
        C = nextByP(Src, " & ");
        E = nextByP(Src, " ^ ");
        cNE = nextByP(Src, " <> ");
        cLT = nextByP(Src, " < ");
        cGT = nextByP(Src, " > ");
        cLE = nextByP(Src, " <= ");
        cGE = nextByP(Src, " >= ");
        cEQ = nextByP(Src, " = ");
        lA = nextByP(Src, " And ");
        lO = nextByP(Src, " Or ");
        lM = nextByP(Src, " Mod ");
        LL = nextByP(Src, " Like ");
        xIs = nextByP(Src, " Is ");
        xLk = nextByP(Src, " Like ");
        P = A; K = 3;
        if (Len(P) > Len(S)) { P = S; K = 3; }
        if (Len(P) > Len(M)) { P = M; K = 3; }
        if (Len(P) > Len(D)) { P = D; K = 3; }
        if (Len(P) > Len(I)) { P = I; K = 3; }
        if (Len(P) > Len(C)) { P = C; K = 3; }
        if (Len(P) > Len(E)) { P = E; K = 3; }
        if (Len(P) > Len(cNE)) { P = cNE; K = 4; }
        if (Len(P) > Len(cLT)) { P = cLT; K = 3; }
        if (Len(P) > Len(cGT)) { P = cGT; K = 3; }
        if (Len(P) > Len(cLE)) { P = cLE; K = 4; }
        if (Len(P) > Len(cGE)) { P = cGE; K = 4; }
        if (Len(P) > Len(cEQ)) { P = cEQ; K = 3; }
        if (Len(P) > Len(lA)) { P = lA; K = 5; }
        if (Len(P) > Len(lO)) { P = lO; K = 4; }
        if (Len(P) > Len(lM)) { P = lM; K = 5; }
        if (Len(P) > Len(LL)) { P = LL; K = 6; }
        if (Len(P) > Len(xLk)) { P = xLk; K = 6; }
        if (Len(P) > Len(xIs)) { P = xIs; K = 4; }
        _NextByOp = P;
        if (Ind <= 1)
        {
            Op = Mid(Src, Len(P) + 1, K);
            _NextByOp = P;
        }
        else
        {
            _NextByOp = _NextByOp(Trim(Mid(Src, Len(P) + 3)), Ind - 1, Op);
        }
        return _NextByOp;
    }
    public static string ReplaceToken(string Src, string OrigToken, string NewToken)
    {
        string _ReplaceToken = "";
        _ReplaceToken = RegExReplace(Src, "([^a-zA-Z_0-9])(" + OrigToken + ")([^a-zA-Z_0-9])", "$1" + NewToken + "$3");
        return _ReplaceToken;
    }
    public static string SplitWord(string Source, int N = 1, string Space = " ", bool TrimResult = true, bool IncludeRest = false)
    {
        string _SplitWord = "";
        // ::::SplitWord
        // :::SUMMARY
        // : Return an indexed word from a string
        // :::DESCRIPTION
        // : Split()s a string based on a space (or other character) and return the word specified by the index.
        // : - Returns __S1 for 1 > N > Count
        // :::PARAMETERS
        // : - Source - The original source string to analyze
        // : - [N] = 1 - The index of the word to return (Default = 1)
        // : - [Space] = __S1 - The character to use as the __S2 (defaults to %20).
        // : - [TrimResult] - Apply Trim() to the result (Default = True)
        // : - [IncludeRest] - Return the rest of the string starting at the indexed word (Default = False).
        // :::EXAMPLE
        // : - SplitWord(__S1, 4) == __S2
        // : - SplitWord(__S1, 4, , , True) == __S2
        // : - SplitWord(__S1, -1, __S2) === __S3
        // :::RETURN
        // :  String
        // :::SEE ALSO
        // : Split, CountWords
        List<string> S = new List<string>();
        int I = 0;
        N = N - 1;
        if (Source == "") return _SplitWord;
        S = new List<string>(Split(Source, Space));
        if (N < 0) N = S.Count + N + 2;
        if (N < 0 || N > S.Count) return _SplitWord;
        if (!IncludeRest)
        {
            _SplitWord = S[N];
        }
        else
        {
            for (I = N; I <= S.Count; I += 1)
            {
                _SplitWord = _SplitWord + IIf(Len(_SplitWord) > 0, Space, "") + S[I];
            }
        }
        if (TrimResult) _SplitWord = Trim(_SplitWord);
        return _SplitWord;
    }
    public static int CountWords(string Source, string Space = " ")
    {
        int _CountWords = 0;
        // ::::CountWords
        // :::SUMMARY
        // : Returns the number of words in a string (determined by <Space> parameter)
        // :::DESCRIPTION
        // : Returns the count of words.
        // :::PARAMETERS
        // : - Source - The original source string to analyze
        // : - [Space] = __S1 - The character to use as the __S2 (defaults to %20).
        // :::EXAMPLE
        // : - CountWords(__S1) == 6
        // : - CountWords(__S1, __S2) == 4
        // :::RETURN
        // :  String
        // :::SEE ALSO
        // : SplitWord
        dynamic L = null;
        // Count actual words.  Blank spaces don't count, before, after, or in the middle.
        // Only a simple split and loop--there may be faster ways...
        foreach (var iterL in new List<string>(Split(Source, Space)))
        {
            L = iterL;
            if (L != "") _CountWords = _CountWords + 1;
        }
        return _CountWords;
    }
    public static dynamic ArrSlice(ref dynamic sourceArray, int fromIndex, int toIndex)
    {
        dynamic _ArrSlice = null;
        int Idx = 0;
        List<dynamic> tempList = new List<dynamic>();
        if (!IsArray(sourceArray)) return _ArrSlice;
        fromIndex = FitRange(0, fromIndex, sourceArray.Count);
        toIndex = FitRange(fromIndex, toIndex, sourceArray.Count);
        for (Idx = fromIndex; Idx <= toIndex; Idx += 1)
        {
            ArrAdd(ref tempList, ref sourceArray(Idx));
        }
        _ArrSlice = tempList;
        return _ArrSlice;
    }
    public static void ArrAdd(ref List<dynamic> Arr, ref dynamic Item)
    {
        int X = 0;
        // TODO: (NOT SUPPORTED): Err.Clear
        // TODO: (NOT SUPPORTED): On Error Resume Next
        X = Arr.Count;
        if (Err().Number != 0)
        {
            Arr = new List<dynamic>() { Item };
            return;
        }
        // TODO: (NOT SUPPORTED): ReDim Preserve Arr(UBound(Arr) + 1)
        Arr[Arr.Count] = Item;
    }
    public static dynamic SubArr(dynamic sourceArray, int fromIndex, int copyLength)
    {
        dynamic _SubArr = null;
        _SubArr = ArrSlice(ref sourceArray, fromIndex, fromIndex + copyLength - 1);
        return _SubArr;
    }
    public static bool InRange(dynamic LBnd, dynamic CHK, dynamic UBnd, bool IncludeBounds = true)
    {
        bool _InRange = false;
        // TODO: (NOT SUPPORTED): On Error Resume Next // because we're doing this as variants..
        if (IncludeBounds)
        {
            _InRange = (CHK >= LBnd) && (CHK <= UBnd);
        }
        else
        {
            _InRange = (CHK > LBnd) && (CHK < UBnd);
        }
        return _InRange;
    }
    public static dynamic FitRange(dynamic LBnd, dynamic CHK, dynamic UBnd)
    {
        dynamic _FitRange = null;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (CHK < LBnd)
        {
            _FitRange = LBnd;
        }
        else if (CHK > UBnd)
        {
            _FitRange = UBnd;
        }
        else
        {
            _FitRange = CHK;
        }
        return _FitRange;
    }
    public static int CodeSectionLoc(string S)
    {
        int _CodeSectionLoc = 0;
        string Token = "Attribute VB_Name";
        int N = 0;
        int K = 0;
        N = InStr(S, Token);
        if (N == 0) return _CodeSectionLoc;
        do
        {
            N = InStr(N, S, vbLf) + 1;
            if (N <= 1) return _CodeSectionLoc;
        } while (Mid(S, N, 10) == "Attribute ");
        _CodeSectionLoc = N;
        return _CodeSectionLoc;
    }
    public static int CodeSectionGlobalEndLoc(string S)
    {
        int _CodeSectionGlobalEndLoc = 0;
        do
        {
            _CodeSectionGlobalEndLoc = _CodeSectionGlobalEndLoc + RegExNPos(Mid(S, _CodeSectionGlobalEndLoc + 1), "([^a-zA-Z0-9_]Function |[^a-zA-Z0-9_]Sub |[^a-zA-Z0-9_]Property )") + 1;
            if (_CodeSectionGlobalEndLoc == 1) { _CodeSectionGlobalEndLoc = Len(S); return _CodeSectionGlobalEndLoc; }
        } while (Mid(S, _CodeSectionGlobalEndLoc - 8, 8) == "Declare ");
        if (_CodeSectionGlobalEndLoc >= 8)
        {
            if (Mid(S, _CodeSectionGlobalEndLoc - 7, 7) == "Friend ") _CodeSectionGlobalEndLoc = _CodeSectionGlobalEndLoc - 7;
            if (Mid(S, _CodeSectionGlobalEndLoc - 7, 7) == "Public ") _CodeSectionGlobalEndLoc = _CodeSectionGlobalEndLoc - 7;
            if (Mid(S, _CodeSectionGlobalEndLoc - 8, 8) == "Private ") _CodeSectionGlobalEndLoc = _CodeSectionGlobalEndLoc - 8;
        }
        _CodeSectionGlobalEndLoc = _CodeSectionGlobalEndLoc - 1;
        return _CodeSectionGlobalEndLoc;
    }
    public static bool isOperator(string S)
    {
        bool _isOperator = false;
        switch (Trim(S))
        {
            case "+":
            case "-":
            case "/":
            case "*":
            case "&":
            case "<>":
            case "<":
            case ">":
            case "<=":
            case ">=":
            case "=":
            case "Mod":
            case "And":
            case "Or":
            case "Xor":
                _isOperator = true;
                break;
            default:
                _isOperator = false;
                break;
        }
        return _isOperator;
    }
    public static void Prg(int Val = -1, int Max = -1, string Cap = "#")
    {
        dynamic L = null;
        bool Found = false;
        foreach (var iterL in Forms)
        {
            L = iterL;
            if (L.Name == "frm") { Found = true; break; }
        }
        if (!Found) return;
        frm.instance.Prg(Val, Max, Cap);
    }
    public static string cVal(ref Collection Coll, string Key, string Def = "")
    {
        string _cVal = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        _cVal = Def;
        _cVal = Coll.Item(LCase(Key));
        return _cVal;
    }
    public static string cValP(ref Collection Coll, string Key, string Def = "")
    {
        string _cValP = "";
        _cValP = P(deQuote(cVal(ref Coll, Key, Def)));
        return _cValP;
    }
    public static string P(string Str)
    {
        string _P = "";
        Str = Replace(Str, "&", "&amp;");
        Str = Replace(Str, "<", "&lt;");
        Str = Replace(Str, ">", "&gt;");
        _P = Str;
        return _P;
    }
    public static string ModuleName(string S)
    {
        string _ModuleName = "";
        int J = 0;
        int K = 0;
        string NameTag = "Attribute VB_Name = \"";
        J = InStr(S, NameTag) + Len(NameTag);
        K = InStr(J, S, "\"") - J;
        _ModuleName = Mid(S, J, K);
        return _ModuleName;
    }
    public static bool IsInCode(string Src, int N)
    {
        bool _IsInCode = false;
        int I = 0;
        string C = "";
        bool Qu = false;
        _IsInCode = false;
        for (I = N; I <= -1; I += 1)
        {
            C = Mid(Src, I, 1);
            if (C == vbCr || C == vbLf)
            {
                _IsInCode = true;
                return _IsInCode;
            }
            else if (C == "\"")
            {
                Qu = !Qu;
            }
            else if (C == "'")
            {
                if (!Qu) return _IsInCode;
            }
        }
        _IsInCode = true;
        return _IsInCode;
    }
    public static string TokenList(string S)
    {
        string _TokenList = "";
        int I = 0;
        int N = 0;
        string T = "";
        N = RegExCount(S, patToken);
        for (I = 0; I <= N - 1; I += 1)
        {
            T = RegExNMatch(S, patToken, I);
            _TokenList = _TokenList + "," + T;
        }
        return _TokenList;
    }
    public static int Random(int Max = 10000)
    {
        int _Random = 0;
        Randomize();
        _Random = ((Rnd() * Max) + 1);
        return _Random;
    }
    public static string Stack(ref string Src, string Val = "##REM##", bool Peek = false)
    {
        string _Stack = "";
        if (Val == "##REM##")
        {
            _Stack = nextBy(Src, ",");
            if (!Peek) Src = Mid(Src, Len(_Stack) + 2);
            _Stack = Replace(_Stack, "\"\"", "\"");
            if (Left(_Stack, 1) == "\"")
            {
                _Stack = Mid(_Stack, 2);
                _Stack = Left(_Stack, Len(_Stack) - 1);
            }
        }
        else
        {
            Src = "\"" + Replace(Val, "\"", "\"\"") + "\"," + Src;
            _Stack = Val;
        }
        return _Stack;
    }
    public static string QuoteXML(string S)
    {
        string _QuoteXML = "";
        _QuoteXML = S;
        _QuoteXML = Replace(S, "\"", "&quot;");
        _QuoteXML = Quote(_QuoteXML);
        return _QuoteXML;
    }
    public static string ReduceString(string Src, string Allowed = "", string Subst = "-", int MaxLen = 0, bool bLCase = true)
    {
        string _ReduceString = "";
        // ::::ReduceString
        // :::SUMMARY
        // : Reduces a string by removing non-allowed characters, optionally replacing them with a substitute.
        // :::DESCRIPTION
        // : Non-allowed characters are removed, and, if supplied, replaced with a substitute.
        // : Substitutes are trimmed from either end, and all duplicated substitutes are remvoed.
        // :
        // : After this process, the string can be given LCase (default) or truncated (not default), if desired.
        // :
        // : This is effectively a slug maker, although it is somewhat adaptable to any cleaning routine.
        // :::PARAMETERS
        // : - Src - Source string to be reduced
        // : - [Allowed] - The list of allowable characters.  Defaults to [A-Za-z0-9]*
        // : - [Subst] - If specified, the character to replace non-allowed characters with (default == __S1)
        // : - [MaxLen] - If passed, truncates longer strings to this length.  Default = 0
        // : - [bLCase] - Convert string to lower case after operation.  Default = True
        // :::EXAMPLE
        // : - ReduceString(__S1) == __S2
        // :::RETURN
        // :  String - The slug generated from the source.
        // :::AUTHOR
        // : Benjamin - 2018.04.28
        // :::SEE ALSO
        // :  ArrangeString, StringNumerals, slug, CleanANI
        int I = 0;
        int N = 0;
        string C = "";
        if (Allowed == "") Allowed = STR_CHR_UCASE + STR_CHR_LCASE + STR_CHR_DIGIT;
        _ReduceString = "";
        N = Len(Src);
        for (I = 1; I <= N; I += 1)
        {
            C = Mid(Src, I, 1);
            _ReduceString = _ReduceString + IIf(IsInStr(Allowed, C), C, Subst);
        }
        if (Subst != "")
        {
            while (IsInStr(_ReduceString, Subst + Subst)) { _ReduceString = Replace(_ReduceString, Subst + Subst, Subst); }
            while (Left(_ReduceString, Len(Subst)) == Subst) { _ReduceString = Mid(_ReduceString, Len(Subst) + 1); }
            while (Right(_ReduceString, Len(Subst)) == Subst) { _ReduceString = Left(_ReduceString, Len(_ReduceString) - Len(Subst)); }
        }
        if (MaxLen > 0) _ReduceString = Left(_ReduceString, MaxLen);
        if (bLCase) _ReduceString = LCase(_ReduceString);
        return _ReduceString;
    }

}
