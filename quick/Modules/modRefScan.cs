using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Strings;
using static modConfig;
using static modControlProperties;
using static modConvertForm;
using static modProjectFiles;
using static modRegEx;
using static modTextFiles;
using static modUtils;
using static modVB6ToCS;
using static VBExtension;



static class modRefScan
{
    public static string OutRes = "";
    public static string cFuncRef_Name = "";
    public static string cFuncRef_Value = "";
    public static string cEnuRef_Name = "";
    public static string cEnumRef_Value = "";
    public static Collection Funcs = null;
    public static Collection LocalFuncs = null;
    private static string RefList(bool KillRef = false)
    {
        string _RefList = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        _RefList = AppContext.BaseDirectory + "\\refs.txt";
        if (KillRef) Kill(_RefList);
        return _RefList;
    }
    public static int FuncsCount(bool vLocal = false)
    {
        int _FuncsCount = 0;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (vLocal)
        {
            _FuncsCount = LocalFuncs.Count;
        }
        else
        {
            _FuncsCount = Funcs.Count;
        }
        return _FuncsCount;
    }
    public static int ScanRefs()
    {
        int _ScanRefs = 0;
        dynamic L = null;
        string T = "";
        string LL = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        OutRes = "";
        _ScanRefs = 0;
        foreach (var iterL in new List<string>(Split(VBPModules(vbpFile), vbCrLf)))
        {
            L = iterL;
            if (L == "") goto SkipMod;
            LL = Replace(L, ".bas", "");
            OutRes = OutRes + vbCrLf + LL + ":" + LL + ":Module:";
            _ScanRefs = _ScanRefs + ScanRefsFile(FilePath(vbpFile) + L);
        SkipMod:;
        }
        foreach (var iterL in new List<string>(Split(VBPForms(vbpFile), vbCrLf)))
        {
            L = iterL;
            L = Replace(L, ".frm", "");
            if (L == "") goto SkipForm;
            T = vbCrLf + L + ":" + L + ":Form:";
            OutRes = OutRes + T;
            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            string S = "";
            int J = 0;
            string Preamble = "";
            string ControlRefs = "";
            S = ReadEntireFile(vbpPath + L + ".frm");
            J = CodeSectionLoc(S);
            Preamble = Left(S, J - 1);
            ControlRefs = FormControls(L, Preamble, false);
            OutRes = OutRes + ControlRefs;
            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            _ScanRefs = _ScanRefs + 1;
        SkipForm:;
        }
        RefList(KillRef:= true);
        WriteFile(RefList(), OutRes);
        OutRes = "";
        return _ScanRefs;
    }
    private static int ScanRefsFile(string FN)
    {
        int _ScanRefsFile = 0;
        string M = "";
        string S = "";
        string L = "";
        dynamic LL = null;
        string F = "";
        string G = "";
        bool Cont = false;
        bool DoCont = false;
        string CurrEnum = "";
        M = FileBaseName(FN);
        S = ReadEntireFile(FN);
        _ScanRefsFile = 0;
        foreach (var iterLL in new List<string>(Split(S, vbCrLf)))
        {
            LL = iterLL;
            DoCont = Right(LL, 1) == "_";
            if (!Cont && !DoCont)
            {
                L = Trim(LL);
                Cont = false;
            }
            else if (Cont && !DoCont)
            {
                L = L + Trim(LL);
                Cont = false;
            }
            else if (!Cont && DoCont)
            {
                L = Trim(Left(LL, Len(LL) - 2));
                Cont = true;
                goto NextLine;
            }
            else if (Cont && DoCont)
            {
                L = L + Trim(Left(LL, Len(LL) - 2));
                Cont = true;
                goto NextLine;
            }
            if (tLMatch(L, "Function ") || tLMatch(L, "Public Function ") || tLMatch(L, "Sub ") || tLMatch(L, "Public Sub ") || false)
            {
                F = Trim(L);
                if (Left(F, 7) == "Public ") F = Mid(F, 8);
                F = Trim(nextBy(F, ":"));
                G = F;
                if (tLMatch(G, "Function ")) G = Mid(G, 10);
                if (tLMatch(G, "Sub ")) G = Mid(G, 5);
                G = nextBy(G, "(");
                F = M + ":" + G + ":Function:" + F;
                OutRes = OutRes + vbCrLf + F;
                _ScanRefsFile = _ScanRefsFile + 1;
            }
            else if (tLMatch(L, "Private Function ") || tLMatch(L, "Private Sub ") || false)
            {
                F = Trim(L);
                F = Trim(nextBy(F, ":"));
                G = F;
                if (tLMatch(G, "Private Function ")) G = Mid(G, 17);
                if (tLMatch(G, "Private Sub ")) G = Mid(G, 12);
                G = nextBy(G, "(");
                F = M + ":" + Trim(M) + "." + Trim(G) + ":Private Function:" + F;
                OutRes = OutRes + vbCrLf + F;
                _ScanRefsFile = _ScanRefsFile + 1;
            }
            else if (tLMatch(L, "Declare ") || tLMatch(L, "Public Decalre "))
            {
                L = LTrim(L);
                if (LMatch(L, "Public ")) L = Mid(L, 8);
                if (LMatch(L, "Declare ")) L = Mid(L, 9);
                G = SplitWord(L);
            }
            else if (tLMatch(L, "Const ") || tLMatch(L, "Public Const ") || tLMatch(L, "Global Const "))
            {
                L = LTrim(L);
                if (LMatch(L, "Public ")) L = Mid(L, 8);
                if (LMatch(L, "Global ")) L = Mid(L, 8);
                if (LMatch(L, "Const ")) L = Mid(L, 7);
                G = SplitWord(L);
            }
            else if (tLMatch(L, "Enum ") || tLMatch(L, "Public Enum "))
            {
                L = LTrim(L);
                if (LMatch(L, "Public ")) L = Mid(L, 8);
                if (LMatch(L, "Enum ")) L = Mid(L, 5);
                CurrEnum = Trim(L);
            }
            else if (tLMatch(L, "End Enum"))
            {
                CurrEnum = "";
            }
            else if (CurrEnum != "")
            {
                G = SplitWord(L);
                F = M + ":" + G + ":Enum:" + CurrEnum + "." + G;
                OutRes = OutRes + vbCrLf + F;
                _ScanRefsFile = _ScanRefsFile + 1;
            }
        NextLine:;
        }
        return _ScanRefsFile;
    }
    public static string ScanRefsFileToString(string FN)
    {
        string _ScanRefsFileToString = "";
        OutRes = "";
        ScanRefsFile(FN);
        _ScanRefsFileToString = OutRes;
        OutRes = "";
        return _ScanRefsFileToString;
    }
    private static void InitFuncs()
    {
        string S = "";
        dynamic L = null;
        if (Dir(RefList()) == "") ScanRefs();
        if (!(Funcs == null)) return;
        S = ReadEntireFile(RefList());
        Funcs = new Collection();
        // TODO: (NOT SUPPORTED): On Error Resume Next
        foreach (var iterL in new List<string>(Split(S, vbCrLf)))
        {
            L = iterL;
            Funcs.Add(L, SplitWord(L, 2, ":"));
        }
        InitLocalFuncs();
    }
    public static void InitLocalFuncs(string S = "")
    {
        // TODO: (NOT SUPPORTED): On Error Resume Next
        dynamic L = null;
        LocalFuncs = new Collection();
        foreach (var iterL in new List<string>(Split(S, vbCrLf)))
        {
            L = iterL;
            LocalFuncs.Add(L, SplitWord(L, 2, ":"));
        }
    }
    public static string FuncRef(string fName)
    {
        string _FuncRef = "";
        if (fName == cFuncRef_Name)
        {
            _FuncRef = cFuncRef_Value;
            return _FuncRef;
        }
        InitFuncs();
        // TODO: (NOT SUPPORTED): On Error Resume Next
        _FuncRef = Funcs(fName);
        if (_FuncRef == "") _FuncRef = LocalFuncs(fName);
        cFuncRef_Name = fName;
        cFuncRef_Value = _FuncRef;
        return _FuncRef;
    }
    public static string FuncRefModule(string fName)
    {
        string _FuncRefModule = "";
        _FuncRefModule = nextBy(FuncRef(fName), ":");
        return _FuncRefModule;
    }
    public static string FuncRefEntity(string fName)
    {
        string _FuncRefEntity = "";
        _FuncRefEntity = nextBy(FuncRef(fName), ":", 3);
        return _FuncRefEntity;
    }
    public static string FuncRefDecl(string fName)
    {
        string _FuncRefDecl = "";
        _FuncRefDecl = nextBy(FuncRef(fName), ":", 4);
        return _FuncRefDecl;
    }
    public static bool IsFuncRef(string fName)
    {
        bool _IsFuncRef = false;
        _IsFuncRef = FuncRef(fName) != "" && FuncRefEntity(fName) == "Function";
        return _IsFuncRef;
    }
    public static bool IsPrivateFuncRef(string Module, string fName)
    {
        bool _IsPrivateFuncRef = false;
        string TName = "";
        TName = Trim(Module) + "." + Trim(fName);
        _IsPrivateFuncRef = FuncRef(TName) != "" && FuncRefEntity(TName) == "Private Function";
        return _IsPrivateFuncRef;
    }
    public static bool IsEnumRef(string fName)
    {
        bool _IsEnumRef = false;
        _IsEnumRef = FuncRef(fName) != "" && FuncRefEntity(fName) == "Enum";
        return _IsEnumRef;
    }
    public static bool IsFormRef(string fName)
    {
        bool _IsFormRef = false;
        string T = "";
        T = SplitWord(fName, 1, ".");
        _IsFormRef = FuncRef(T) != "" && FuncRefEntity(T) == "Form";
        return _IsFormRef;
    }
    public static bool IsModuleRef(string fName)
    {
        bool _IsModuleRef = false;
        string T = "";
        T = SplitWord(fName, 1, ".");
        _IsModuleRef = FuncRef(T) != "" && FuncRefEntity(T) == "Module";
        return _IsModuleRef;
    }
    public static bool IsControlRef(string Src, string FormName = "")
    {
        bool _IsControlRef = false;
        string Tok = "";
        string Tok2 = "";
        string FTok = "";
        string TTok = "";
        Tok = RegExNMatch(Src, patToken);
        Tok2 = RegExNMatch(Src, patToken, 1);
        TTok = Tok + "." + Tok2;
        FTok = FormName + "." + Tok;
        // If IsInStr(Src, __S1) Then Stop
        if (FuncRef(TTok) != "" && FuncRefEntity(TTok) == "Control" || FuncRef(FTok) != "" && FuncRefEntity(FTok) == "Control")
        {
            _IsControlRef = true;
        }
        return _IsControlRef;
    }
    public static string FormControlRefDeclType(string Src, string FormName = "")
    {
        string _FormControlRefDeclType = "";
        string Tok = "";
        string Tok2 = "";
        string FTok = "";
        string TTok = "";
        Tok = RegExNMatch(Src, patToken);
        Tok2 = RegExNMatch(Src, patToken, 1);
        TTok = Tok + "." + Tok2;
        FTok = FormName + "." + Tok;
        // If IsInStr(Src, __S1) Then Stop
        if (FuncRef(TTok) != "" && FuncRefEntity(TTok) == "Control")
        {
            _FormControlRefDeclType = FuncRefDecl(TTok);
        }
        else if (FuncRef(FTok) != "" && FuncRefEntity(FTok) == "Control")
        {
            _FormControlRefDeclType = FuncRefDecl(FTok);
        }
        return _FormControlRefDeclType;
    }
    public static string FuncRefDeclTyp(string fName)
    {
        string _FuncRefDeclTyp = "";
        _FuncRefDeclTyp = SplitWord(FuncRefDecl(fName), 1);
        return _FuncRefDeclTyp;
    }
    public static string FuncRefDeclRet(string fName)
    {
        string _FuncRefDeclRet = "";
        _FuncRefDeclRet = FuncRefDecl(fName);
        _FuncRefDeclRet = Trim(Mid(_FuncRefDeclRet, InStrRev(_FuncRefDeclRet, " ")));
        if (Right(_FuncRefDeclRet, 1) == ")" && Right(_FuncRefDeclRet, 2) != "()") _FuncRefDeclRet = "";
        return _FuncRefDeclRet;
    }
    public static string FuncRefDeclArgs(string fName)
    {
        string _FuncRefDeclArgs = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        _FuncRefDeclArgs = FuncRefDecl(fName);
        _FuncRefDeclArgs = Mid(_FuncRefDeclArgs, InStr(_FuncRefDeclArgs, "(") + 1);
        _FuncRefDeclArgs = Left(_FuncRefDeclArgs, InStrRev(_FuncRefDeclArgs, ")") - 1);
        _FuncRefDeclArgs = Trim(_FuncRefDeclArgs);
        return _FuncRefDeclArgs;
    }
    public static string FuncRefDeclArgN(string fName, int N)
    {
        string _FuncRefDeclArgN = "";
        string F = "";
        F = FuncRefDeclArgs(fName);
        _FuncRefDeclArgN = nextBy(F, ", ", N);
        return _FuncRefDeclArgN;
    }
    public static int FuncRefDeclArgCnt(string fName)
    {
        int _FuncRefDeclArgCnt = 0;
        string F = "";
        string K = "";
        F = FuncRefDeclArgs(fName);
        _FuncRefDeclArgCnt = 0;
        do
        {
            K = nextBy(F, ", ", _FuncRefDeclArgCnt + 1);
            if (K == "") return _FuncRefDeclArgCnt;
            _FuncRefDeclArgCnt = _FuncRefDeclArgCnt + 1;
        } while (true);
        return _FuncRefDeclArgCnt;
    }
    public static string FuncRefArgType(string fName, int N)
    {
        string _FuncRefArgType = "";
        _FuncRefArgType = FuncRefDeclArgN(fName, N);
        if (_FuncRefArgType == "") return _FuncRefArgType;
        _FuncRefArgType = SplitWord(_FuncRefArgType, 2, " As ");
        return _FuncRefArgType;
    }
    public static bool FuncRefArgByRef(string fName, int N)
    {
        bool _FuncRefArgByRef = false;
        _FuncRefArgByRef = !IsInStr(FuncRefDeclArgN(fName, N), "ByVal ");
        return _FuncRefArgByRef;
    }
    public static bool FuncRefArgOptional(string fName, int N)
    {
        bool _FuncRefArgOptional = false;
        _FuncRefArgOptional = IsInStr(FuncRefDeclArgN(fName, N), "Optional ");
        return _FuncRefArgOptional;
    }
    public static string FuncRefArgDefault(string fName, int N)
    {
        string _FuncRefArgDefault = "";
        string aTyp = "";
        if (!FuncRefArgOptional(fName, N)) return _FuncRefArgDefault;
        _FuncRefArgDefault = SplitWord(FuncRefDeclArgN(fName, N), 2, " = ", true, true);
        if (_FuncRefArgDefault == "") _FuncRefArgDefault = ConvertDefaultDefault(FuncRefArgType(fName, N));
        return _FuncRefArgDefault;
    }
    public static string EnumRefRepl(string EName)
    {
        string _EnumRefRepl = "";
        _EnumRefRepl = FuncRefDecl(EName);
        return _EnumRefRepl;
    }
    public static string FormRefRepl(string fName)
    {
        string _FormRefRepl = "";
        string T = "";
        string U = "";
        T = SplitWord(fName, 1, ".");
        U = FuncRefModule(T) + ".instance";
        _FormRefRepl = Replace(fName, T, U);
        return _FormRefRepl;
    }
    public static string FormControlRepl(string Src, string FormName = "")
    {
        string _FormControlRepl = "";
        string Tok = "";
        string Tok2 = "";
        string Tok3 = "";
        string F = "";
        string V = "";
        Tok = RegExNMatch(Src, patToken);
        Tok2 = RegExNMatch(Src, patToken, 1);
        Tok3 = RegExNMatch(Src, patToken, 2);
        // If IsInStr(Tok, __S1) Then Stop
        // If IsInStr(Src, __S1) Then Stop
        if (!IsFormRef(Tok))
        {
            F = Tok;
            V = ConvertControlProperty(F, Tok2, FuncRefDecl(FormName + "." + Tok));
            if (Tok2 != "")
            {
                _FormControlRepl = Replace(Src, Tok2, V);
            }
            else
            {
                _FormControlRepl = Src + "." + V;
            }
        }
        else
        {
            F = Tok + "." + Tok2;
            V = ConvertControlProperty(F, Tok3, FuncRefDecl(Tok + "." + Tok2));
            if (Tok3 != "")
            {
                _FormControlRepl = Replace(Src, Tok3, V);
            }
            else
            {
                _FormControlRepl = Src + "." + V;
            }
        }
        return _FormControlRepl;
    }

}
