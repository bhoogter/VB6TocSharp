using System;
using System.Collections.Generic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modConfig;
using static modConvertForm;
using static modOrigConvert;
using static modProjectFiles;
using static modQuickConvert;
using static modRefScan;
using static modSupportFiles;
using static modTextFiles;
using static modUsingEverything;
using static modUtils;
using static VBExtension;



static class modConvert
{
    public const string WithMark = "_WithVar_";
    public static int WithLevel = 0;
    public static int MaxWithLevel = 0;
    public static string WithVars = "";
    public static string WithTypes = "";
    public static string WithAssign = "";
    public static string FormName = "";
    public static string CurrentModule = "";
    public static string CurrSub = "";
    public const string CONVERTER_VERSION_1 = "v1";
    public const string CONVERTER_VERSION_2 = "v2";
    public const string CONVERTER_VERSION_DEFAULT = CONVERTER_VERSION_2;
    public static bool QuickConvertProject()
    {
        bool _QuickConvertProject = false;
        _QuickConvertProject = ConvertProject(vbpFile, CONVERTER_VERSION_2);
        return _QuickConvertProject;
    }
    public static bool QuickConvert()
    {
        bool _QuickConvert = false;
        _QuickConvert = ConvertFile("modQuickConvert.bas", false, CONVERTER_VERSION_2);
        return _QuickConvert;
    }
    public static bool ConvertProject(string vbpFile = "", string ConverterVersion = CONVERTER_VERSION_DEFAULT)
    {
        bool _ConvertProject = false;
        if (vbpFile == "") vbpFile = modConfig.vbpFile;
        Prg(0, 1, "Preparing...");
        ScanRefs();
        CreateProjectFile(vbpFile);
        CreateProjectSupportFiles();
        ConvertFileList(FilePath(vbpFile), VBPModules(vbpFile) + vbCrLf + VBPClasses(vbpFile) + vbCrLf + VBPForms(vbpFile), vbCrLf, ConverterVersion);
        _ConvertProject = true;
        return _ConvertProject;
    }
    public static bool ConvertFileList(string Path, string List, string Sep = vbCrLf, string ConverterVersion = CONVERTER_VERSION_DEFAULT)
    {
        bool _ConvertFileList = false;
        dynamic L = null;
        int V = 0;
        int N = 0;
        V = StrCnt(List, Sep) + 1;
        Prg(0, V, N + "/" + V + "...");
        foreach (var iterL in new List<string>(Split(List, Sep)))
        {
            L = iterL;
            N = N + 1;
            if (L == "") goto NextItem;
            if (L == "modFunctionList.bas") goto NextItem;
            ConvertFile(Path + L, false, ConverterVersion);
        NextItem:;
            Prg(N, -1 , N + "/" + V + ": " + L);
            DoEvents();
        }
        Prg();
        return _ConvertFileList;
    }
    public static bool ConvertFile(string SomeFile, bool UIOnly = false, string ConverterVersion = CONVERTER_VERSION_DEFAULT)
    {
        bool _ConvertFile = false;
        if (!IsInStr(SomeFile, "\\")) SomeFile = vbpPath + SomeFile;
        CurrentModule = "";
        switch (LCase(FileExt(SomeFile)))
        {
            case ".bas":
                _ConvertFile = ConvertModule(SomeFile, ConverterVersion);
                break;
            case ".cls":
                _ConvertFile = ConvertClass(SomeFile, ConverterVersion);
                break;
            case ".frm":
                FormName = FileBaseName(SomeFile); _ConvertFile = ConvertForm(SomeFile, UIOnly, ConverterVersion);
                // Case __S1: ConvertModule  someFile
                break;
            default:
                MsgBox("UNKNOWN VB TYPE: " + SomeFile); return _ConvertFile;
                break;
        }
        FormName = "";
        _ConvertFile = true;
        return _ConvertFile;
    }
    public static bool ConvertForm(string frmFile, bool UIOnly = false, string ConverterVersion = CONVERTER_VERSION_DEFAULT)
    {
        bool _ConvertForm = false;
        string S = "";
        int J = 0;
        string Preamble = "";
        string Code = "";
        string Globals = "";
        string Functions = "";
        string X = "";
        string fName = "";
        string F = "";
        if (!FileExists(frmFile))
        {
            MsgBox("File not found in ConvertForm: " + frmFile);
            return _ConvertForm;
        }
        S = ReadEntireFile(frmFile);
        fName = ModuleName(S);
        CurrentModule = fName;
        F = fName + ".xaml.cs";
        if (IsConverted(F, frmFile)) { Console.WriteLine("Form Already Converted: " + F); return _ConvertForm; }
        J = CodeSectionLoc(S);
        Preamble = Left(S, J - 1);
        Code = Mid(S, J);
        X = ConvertFormUi(Preamble, Code);
        F = fName + ".xaml";
        WriteOut(F, X, frmFile);
        if (UIOnly) return _ConvertForm;
        string ConvertedCode = "";
        if (ConverterVersion == CONVERTER_VERSION_2)
        {
            ConvertedCode = "";
            string ControlArrays = "";
            dynamic VV = null;
            ControlArrays = Replace(Replace(Replace(modConvertForm.FormControlArrays, "][", ";"), "[", ""), "]", "");
            foreach (var iterVV in new List<string>(Split(ControlArrays, ";")))
            {
                VV = iterVV;
                List<string> ControlArrayParts = new List<string>();
                ControlArrayParts = new List<string>(Split(VV, ","));
                ConvertedCode = ConvertedCode + "public List<" + ControlArrayParts[1] + "> " + ControlArrayParts[0] + " { get => VBExtension.controlArray<" + ControlArrayParts[1] + ">(this, \"" + ControlArrayParts[0] + "\"); }" + vbCrLf2;
                // ConvertedCode = ConvertedCode & __S1 & ControlArrayParts(1) & __S2 & ControlArrayParts(0) & __S3 & ControlArrayParts(1) & __S4 & ControlArrayParts(1) & __S5 & vbCrLf
                // ConvertedCode = ConvertedCode & __S1 & ControlArrayParts(1) & __S2 & ControlArrayParts(0) & __S3 & ControlArrayParts(0) & __S4 & vbCrLf2
            }
            ConvertedCode = ConvertedCode + QuickConvertFile(frmFile);
        }
        else
        {
            J = CodeSectionGlobalEndLoc(Code);
            Globals = ConvertGlobals(Left(Code, J));
            InitLocalFuncs(FormControls(fName, Preamble) + ScanRefsFileToString(frmFile));
            Functions = ConvertCodeSegment(Mid(Code, J));
            ConvertedCode = Globals + vbCrLf2 + Functions;
        }
        X = "";
        X = X + UsingEverything(fName) + vbCrLf;
        X = X + vbCrLf;
        X = X + "namespace " + AssemblyName() + ".Forms" + vbCrLf;
        X = X + "{" + vbCrLf;
        X = X + "public partial class " + fName + " : Window {" + vbCrLf;
        X = X + "  private static " + fName + " _instance;" + vbCrLf;
        X = X + "  public static " + fName + " instance { set { _instance = null; } get { return _instance ?? (_instance = new " + fName + "()); }}";
        X = X + "  public static void Load() { if (_instance == null) { dynamic A = " + fName + ".instance; } }";
        X = X + "  public static void Unload() { if (_instance != null) instance.Close(); _instance = null; }";
        X = X + "  public " + fName + "() { InitializeComponent(); }" + vbCrLf;
        X = X + vbCrLf;
        X = X + vbCrLf;
        X = X + ConvertedCode;
        X = X + vbCrLf + "}";
        X = X + vbCrLf + "}";
        X = deWS(X);
        F = fName + ".xaml.cs";
        WriteOut(F, X, frmFile);
        return _ConvertForm;
    }
    public static bool ConvertModule(string basFile, string ConverterVersion = CONVERTER_VERSION_DEFAULT)
    {
        bool _ConvertModule = false;
        string S = "";
        int J = 0;
        string Code = "";
        string Globals = "";
        string Functions = "";
        string F = "";
        string X = "";
        string fName = "";
        if (!FileExists(basFile))
        {
            MsgBox("File not found in ConvertModule: " + basFile);
            return _ConvertModule;
        }
        S = ReadEntireFile(basFile);
        fName = ModuleName(S);
        CurrentModule = fName;
        F = fName + ".cs";
        if (IsConverted(F, basFile)) { Console.WriteLine("Module Already Converted: " + F); return _ConvertModule; }
        fName = ModuleName(S);
        Code = Mid(S, CodeSectionLoc(S));
        string UserCode = "";
        if (ConverterVersion == CONVERTER_VERSION_2)
        {
            UserCode = QuickConvertFile(basFile);
        }
        else
        {
            J = CodeSectionGlobalEndLoc(Code);
            Globals = ConvertGlobals(Left(Code, J - 1), true);
            Functions = ConvertCodeSegment(Mid(Code, J), true);
            UserCode = nlTrim(Globals + vbCrLf + vbCrLf + Functions);
            UserCode = deWS(UserCode);
        }
        X = "";
        X = X + UsingEverything(fName) + vbCrLf;
        X = X + vbCrLf;
        X = X + "static class " + fName + " {" + vbCrLf;
        X = X + UserCode;
        X = X + vbCrLf + "}";
        WriteOut(F, X, basFile);
        return _ConvertModule;
    }
    public static bool ConvertClass(string clsFile, string ConverterVersion = CONVERTER_VERSION_DEFAULT)
    {
        bool _ConvertClass = false;
        string S = "";
        int J = 0;
        string Code = "";
        string Globals = "";
        string Functions = "";
        string F = "";
        string X = "";
        string fName = "";
        string cName = "";
        if (!FileExists(clsFile))
        {
            MsgBox("File not found in ConvertModule: " + clsFile);
            return _ConvertClass;
        }
        S = ReadEntireFile(clsFile);
        fName = ModuleName(S);
        CurrentModule = fName;
        F = fName + ".cs";
        if (IsConverted(F, clsFile)) { Console.WriteLine("Class Already Converted: " + F); return _ConvertClass; }
        string UserCode = "";
        if (ConverterVersion == CONVERTER_VERSION_2)
        {
            UserCode = QuickConvertFile(clsFile);
        }
        else
        {
            Code = Mid(S, CodeSectionLoc(S));
            J = CodeSectionGlobalEndLoc(Code);
            Globals = ConvertGlobals(Left(Code, J - 1));
            Functions = ConvertCodeSegment(Mid(Code, J));
            UserCode = deWS(Globals + vbCrLf + vbCrLf + Functions);
        }
        X = "";
        X = X + UsingEverything(fName) + vbCrLf;
        X = X + vbCrLf;
        X = X + "public class " + fName + " {" + vbCrLf;
        X = X + UserCode;
        X = X + vbCrLf + "}";
        F = fName + ".cs";
        WriteOut(F, X, clsFile);
        return _ConvertClass;
    }

}
