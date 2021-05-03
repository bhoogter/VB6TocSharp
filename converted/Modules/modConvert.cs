using VB6 = Microsoft.VisualBasic.Compatibility.VB6;
using System.Runtime.InteropServices;
using static VBExtension;
using static VBConstants;
using Microsoft.VisualBasic;
using System;
using System.Windows;
using System.Windows.Controls;
using static System.DateTime;
using static System.Math;
using static Microsoft.VisualBasic.Globals;
using static Microsoft.VisualBasic.Collection;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Conversion;
using static Microsoft.VisualBasic.DateAndTime;
using static Microsoft.VisualBasic.ErrObject;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Financial;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static Microsoft.VisualBasic.VBMath;
using System.Collections.Generic;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ColorConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.DrawStyleConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.FillStyleConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.GlobalModule;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.Printer;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterCollection;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterObjectConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ScaleModeConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.SystemColorConstants;
using ADODB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using VB2CS.Forms;
using static modUtils;
using static modConvert;
using static modProjectFiles;
using static modTextFiles;
using static modRegEx;
using static frmTest;
using static modConvertForm;
using static modSubTracking;
using static modVB6ToCS;
using static modUsingEverything;
using static modSupportFiles;
using static modConfig;
using static modRefScan;
using static modConvertUtils;
using static modControlProperties;
using static modProjectSpecific;
using static modINI;
using static modLinter;
using static modGit;
using static modDirStack;
using static modShell;
using static VB2CS.Forms.frm;
using static VB2CS.Forms.frmConfig;


static class modConvert {
// Option Explicit
public const dynamic WithMark = "_WithVar_";
private static int WithLevel = 0;
private static int MaxWithLevel = 0;
private static string WithVars = "";
private static string WithTypes = "";
private static string WithAssign = "";
private static string FormName = "";
private static string CurrSub = "";


public static dynamic ConvertProject(string vbpFile) {
  dynamic ConvertProject = null;
  Prg(0, 1, "Preparing...");
  ScanRefs();
  CreateProjectFile(vbpFile);
  CreateProjectSupportFiles();
  ConvertFileList(FilePath(vbpFile), VBPModules(vbpFile) + vbCrLf + VBPClasses(vbpFile) + vbCrLf + VBPForms(vbpFile)); //& vbCrLf & VBPUserControls(vbpFile)
  MsgBox("Complete.");
  return ConvertProject;
}

public static bool ConvertFileList(string Path, string List, string Sep= vbCrLf) {
  bool ConvertFileList = false;
  dynamic L = null;
  int V = 0;
  int N = 0;

  V = StrCnt(List, Sep) + 1;
  Prg(0, V, N + "/" + V + "...");
  foreach(var iterL in Split(List, Sep)) {
L = iterL;
    N = N + 1;
    if (L == "") {
goto NextItem;
    }

    if (L == "modFunctionList.bas") {
goto NextItem;
    }

    ConvertFile(Path + L);

NextItem:;
    Prg(N);
    DoEvents();
  }
  Prg();
  return ConvertFileList;
}

public static bool ConvertFile(string someFile, bool UIOnly= false) {
  bool ConvertFile = false;
  if (!IsInStr(someFile, "\\")) {
    someFile = vbpPath + someFile;
  }
  switch(LCase(FileExt(someFile))) {
    case ".bas":
      ConvertFile = ConvertModule(someFile);
      break;
    case ".cls":
      ConvertFile = ConvertClass(someFile);
      break;
    case ".frm":
      FormName = FileBaseName(someFile);
      ConvertFile = ConvertForm(someFile, UIOnly);
//      Case ".ctl": ConvertModule  someFile
      break;
    default:
      MsgBox("UNKNOWN VB TYPE: " + someFile);
      return ConvertFile;

break;
}
  FormName = "";
  ConvertFile = true;
  return ConvertFile;
}

public static bool ConvertForm(string frmFile, bool UIOnly= false) {
  bool ConvertForm = false;
  string S = "";
  int J = 0;
  string Preamble = "";
  string Code = "";
  string Globals = "";
  string Functions = "";

  string X = "";
  string fName = "";

  string F = "";

  if (!FileExists(frmFile)) {
    MsgBox("File not found in ConvertForm: " + frmFile);
    return ConvertForm;

  }

  S = ReadEntireFile(frmFile);
  fName = ModuleName(S);
  F = fName + ".xaml.cs";
  if (IsConverted(F, frmFile)) {
    Console.WriteLine("Form Already Converted: " + F);
    return ConvertForm;

  }

  J = CodeSectionLoc(S);
  Preamble = Left(S, J - 1);
  Code = Mid(S, J);

  X = ConvertFormUi(Preamble, Code);
  F = fName + ".xaml";
  WriteOut(F, X, frmFile);
  if (UIOnly) {
    return ConvertForm;

  }

  J = CodeSectionGlobalEndLoc(Code);
  Globals = ConvertGlobals(Left(Code, J));
  InitLocalFuncs(FormControls(fName, Preamble) + ScanRefsFileToString(frmFile));
  Functions = ConvertCodeSegment(Mid(Code, J));

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
  X = X + Globals + vbCrLf + vbCrLf + Functions;
  X = X + vbCrLf + "}";
  X = X + vbCrLf + "}";

  X = deWS(X);

  F = fName + ".xaml.cs";
  WriteOut(F, X, frmFile);
  return ConvertForm;
}

public static dynamic ConvertModule(string basFile) {
  dynamic ConvertModule = null;
  string S = "";
  int J = 0;
  string Code = "";
  string Globals = "";
  string Functions = "";

  string F = "";
  string X = "";
  string fName = "";

  if (!FileExists(basFile)) {
    MsgBox("File not found in ConvertModule: " + basFile);
    return ConvertModule;

  }
  S = ReadEntireFile(basFile);
  fName = ModuleName(S);
  F = fName + ".cs";
  if (IsConverted(F, basFile)) {
    Console.WriteLine("Module Already Converted: " + F);
    return ConvertModule;

  }

  fName = ModuleName(S);
  Code = Mid(S, CodeSectionLoc(S));

  J = CodeSectionGlobalEndLoc(Code);
  Globals = ConvertGlobals(Left(Code, J - 1), true);
  Functions = ConvertCodeSegment(Mid(Code, J), true);

  X = "";
  X = X + UsingEverything(fName) + vbCrLf;
  X = X + vbCrLf;
  X = X + "static class " + fName + " {" + vbCrLf;
  X = X + nlTrim(Globals + vbCrLf + vbCrLf + Functions);
  X = X + vbCrLf + "}";

  X = deWS(X);

  WriteOut(F, X, basFile);
  return ConvertModule;
}

public static dynamic ConvertClass(string clsFile) {
  dynamic ConvertClass = null;
  string S = "";
  int J = 0;
  string Code = "";
  string Globals = "";
  string Functions = "";

  string F = "";
  string X = "";
  string fName = "";

  string cName = "";

  if (!FileExists(clsFile)) {
    MsgBox("File not found in ConvertModule: " + clsFile);
    return ConvertClass;

  }
  S = ReadEntireFile(clsFile);
  fName = ModuleName(S);
  F = fName + ".cs";
  if (IsConverted(F, clsFile)) {
    Console.WriteLine("Class Already Converted: " + F);
    return ConvertClass;

  }

  Code = Mid(S, CodeSectionLoc(S));

  J = CodeSectionGlobalEndLoc(Code);
  Globals = ConvertGlobals(Left(Code, J - 1));
  Functions = ConvertCodeSegment(Mid(Code, J));

  X = "";
  X = X + UsingEverything(fName) + vbCrLf;
  X = X + vbCrLf;
  X = X + "public class " + fName + " {" + vbCrLf;
  X = X + Globals + vbCrLf + vbCrLf + Functions;
  X = X + vbCrLf + "}";

  X = deWS(X);

  F = fName + ".cs";
  WriteOut(F, X, clsFile);
  return ConvertClass;
}

public static string GetMultiLineSpace(string Prv, string Nxt) {
  string GetMultiLineSpace = "";
  string pC = "";
  string nC = "";

  GetMultiLineSpace = " ";
  pC = Right(Prv, 1);
  nC = Left(Nxt, 1);
  if (nC == "(") {
    GetMultiLineSpace = "";
  }
  return GetMultiLineSpace;
}

public static dynamic SanitizeCode(string Str) {
  dynamic SanitizeCode = null;
  const dynamic NamedParamSrc = ":=";
  const dynamic NamedParamTok = "###NAMED-PARAMETER###";
  dynamic Sp = null;
  dynamic L = null;

  string F = "";

  string R = "";
  string N = "";

  string Building = "";

  bool FinishSplitIf = false;


  R = "";
  N = vbCrLf;
  Sp = Split(Str, vbCrLf);
  Building = "";


  foreach(var iterL in Sp) {
L = iterL;
//If IsInStr(L, "POEDIFolder") Then Stop
//If IsInStr(L, "Set objSourceArNo = New_CDbTypeAhead") Then Stop
    if (Right(L, 1) == "_") {
      string C = "";

      C = Trim(Left(L, Len(L) - 1));
      Building = Building + GetMultiLineSpace(Building, C) + C;
goto NextLine;
    }
    if (Building != "") {
      L = Building + GetMultiLineSpace(Building, Trim(L)) + Trim(L);
      Building = "";
    }

//    If IsInStr(L, "'") Then Stop
    L = DeComment(L);
    L = DeString(L);
//If IsInStr(L, "CustRec <> 0") Then Stop

    FinishSplitIf = false;
    if (tLeft(L, 3) == "If " && Right(RTrim(L), 5) != " Then") {
      FinishSplitIf = true;
      F = nextBy(L, " Then ") + " Then";
      R = R + N + F;
      L = Mid(L, Len(F) + 2);
      if (nextBy(L, " Else ", 2) != "") {
        R = R + SanitizeCode[nextBy(L, " Else ", 1)];
        R = R + N + "Else";
        L = nextBy(L, "Else ", 2);
      }
    }

    if (nextBy(L, ":") != L) {
      if (RegExTest(Trim(L), "^[a-zA-Z_][a-zA-Z_0-9]*:$")) { // Goto Label
        R = R + N + ReComment(L);
      } else {
        do {
          L = Replace(L, NamedParamSrc, NamedParamTok);
          F = nextBy(L, ":");
          F = Replace(F, NamedParamTok, NamedParamSrc);
          R = R + N + ReComment(F, true);
          L = Replace(L, NamedParamTok, NamedParamSrc);
          if (F == L) {
            break;
          }
          L = Trim(Mid(L, Len(F) + 2));
          R = R + SanitizeCode[L];

        } while(!(false));
      }
    } else {
      R = R + N + ReComment(L, true);
    }

    if (FinishSplitIf) {
      R = R + N + "End If";
    }
NextLine:;
  }

  SanitizeCode = R;
  return SanitizeCode;
}

public static string ConvertCodeSegment(string S, bool asModule= false) {
  string ConvertCodeSegment = "";
  string P = "";
  int N = 0;

  string F = "";
  int T = 0;
  int E = 0;
  string K = "";
  int X = 0;

  string Pre = "";
  string Body = "";

  string R = "";


  ClearProperties();

  InitDeString();
//WriteFile "C:\Users\benja\Desktop\code.txt", S, True
  S = SanitizeCode(S);
//WriteFile "C:\Users\benja\Desktop\sani.txt", S, True
  do {
    P() = "(Public |Private |)(Friend |)(Function |Sub |Property Get |Property Let |Property Set )" + patToken + "[ ]*\\(";
    N = -1;
    do {
      N = N + 1;
      F = RegExNMatch(S, P(), N);
      T = RegExNPos(S, P(), N);
    } while(!(!IsInCode(S, T) && F != ""));
    if (F == "") {
      break;
    }

    if (IsInStr(F, " Function ")) {
      K = "End Function";
    }
    if (IsInStr(F, " Sub ")) {
      K = "End Sub";
    }
    if (IsInStr(F, " Property ")) {
      K = "End Property";
    }
    N = -1;
    do {
      N = N + 1;
      E = RegExNPos(Mid(S, T), K, N) + Len(K) + T;
    } while(!(!IsInCode(S, E) && E != 0));

    if (T > 1) {
      Pre = nlTrim(Left(S, T - 1));
    } else {
      Pre = "";
    }
    while(!(Mid(S, E, 1) == vbCr || Mid(S, E, 1) == vbLf || Mid(S, E, 1) == "")) {
      E = E + 1;
    }
    Body = nlTrim(Mid(S, T, E - T));

    S = nlTrim(Mid(S, E + 1));

    R = R + CommentBlock(Pre) + ConvertSub(Body, asModule) + vbCrLf;
  } while(!(true));

  R = ReadOutProperties(asModule) + vbCrLf2 + R;

  R = ReString(R, true);

  ConvertCodeSegment = R;
  return ConvertCodeSegment;
}

public static string CommentBlock(string Str) {
  string CommentBlock = "";
  string S = "";

  if (nlTrim(Str) == "") {
    return CommentBlock;

  }
  S = "";
  S = S + "/*" + vbCrLf;
  S = S + Replace(Str, "*/", "* /") + vbCrLf;
  S = S + "*/" + vbCrLf;
  CommentBlock = S;
  return CommentBlock;
}

public static string ConvertDeclare(string S, int Ind, bool isGlobal= false, bool asModule= false) {
  string ConvertDeclare = "";
  dynamic Sp = null;
  dynamic L = null;
  string SS = "";

  bool asPrivate = false;

  string pName = "";
  string pType = "";
  bool pWithEvents = false;

  string Res = "";

  string ArraySpec = "";
  bool isArr = false;
  int aMax = 0;
  int aMin = 0;
  string aTodo = "";

  Res = "";

  SS = S;

  if (tLeft(S, 7) == "Public ") {
    S = tMid(S, 8);
  }
  if (tLeft(S, 4) == "Dim ") {
    S = Mid(Trim(S), 5);
    asPrivate = true;
  }
  if (tLeft(S, 8) == "Private ") {
    S = tMid(S, 9);
    asPrivate = true;
  }

//  If IsInStr(S, "aMin") Then Stop
  Sp = Split(S, ",");
  foreach(var iterL in Sp) {
L = iterL;
    L = Trim(L);
    if (LMatch(L, "WithEvents ")) {
      L = Trim(tMid(L, 12));
      Res = Res + "// TODO: WithEvents not supported on " + RegExNMatch(L, patToken) + vbCrLf;
    }
    pName = RegExNMatch(L, patToken);
    L = Trim(tMid(L, Len(pName) + 1));
    if (isGlobal) {
      Res = Res + IIf(asPrivate, "private ", "public ");
    }
    if (asModule) {
      Res = Res + "static ";
    }
    if (tLeft(L, 1) == "(") {
      isArr = true;
      ArraySpec = nextBy(Mid(L, 2), ")");
      if (ArraySpec == "") {
        aMin = -1;
        aMax = -1;
        L = Trim(tMid(L, 3));
      } else {
        L = Trim(tMid(L, Len(ArraySpec) + 3));
        aMin = 0;
        aMax = Val(SplitWord(ArraySpec));
        ArraySpec = Trim(tMid(ArraySpec, Len(aMax) + 1));
        if (tLeft(ArraySpec, 3) == "To ") {
          aMin = aMax;
          aMax = Val(tMid(ArraySpec, 4));
        }
      }
    }

    bool AsNew = false;

    AsNew = false;
    if (SplitWord(L, 1) == "As") {
      pType = SplitWord(L, 2);
      if (pType == "New") {
        pType = SplitWord(L, 3);
        AsNew = true;
      }
    } else {
      pType = "Variant";
    }

    if (!isArr) {
      Res = Res + sSpace(Ind) + ConvertDataType(pType) + " " + pName;
      Res = Res + " = ";
      if (AsNew) {
        Res = Res + "new ";
        Res = Res + ConvertDataType(pType);
        Res = Res + "()";
      } else {
        Res = Res + ConvertDefaultDefault(pType);
      }
      Res = Res + ";" + vbCrLf;
    } else {
      aTodo = IIf(aMin == 0, "", " // TODO - Specified Minimum Array Boundary Not Supported: " + SS);
      if (!IsNumeric(aMax)) {
        Res = Res + sSpace(Ind) + "List<" + ConvertDataType(pType) + "> " + pName + " = new List<" + ConvertDataType(pType) + "> (new " + ConvertDataType(pType) + "[(" + aMax + " + 1)]);  // TODO: Confirm Array Size By Token" + aTodo + vbCrLf;
      } else if (Val(aMax) == -1) {
        Res = Res + sSpace(Ind) + "List<" + ConvertDataType(pType) + "> " + pName + " = new List<" + ConvertDataType(pType) + "> {};" + aTodo + vbCrLf;
      } else {
        Res = Res + sSpace(Ind) + "List<" + ConvertDataType(pType) + "> " + pName + " = new List<" + ConvertDataType(pType) + "> (new " + ConvertDataType(pType) + "[" + (Val(aMax) + 1) + "]);" + aTodo + vbCrLf;
      }
    }

    SubParamDecl(pName, pType, IIf(isArr, "" + aMax, ""), false, false);
  }

  ConvertDeclare = Res;
  return ConvertDeclare;
}

public static string ConvertAPIDef(string S) {
  string ConvertAPIDef = "";
//Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
//[DllImport("User32.dll")]
//public static extern int MessageBox(int h, string m, string c, int type);
  bool isPrivate = false;
  bool isSub = false;

  string AName = "";

  string aLib = "";

  string aAlias = "";

  string aArgs = "";

  string aReturn = "";

  string tArg = "";
  bool has = false;

  if (tLeft(S, 8) == "Private ") {
    S = tMid(S, 9);
    isPrivate = true;
  }
  if (tLeft(S, 7) == "Public ") {
    S = tMid(S, 8);
  }
  if (tLeft(S, 8) == "Declare ") {
    S = tMid(S, 9);
  }
  if (tLeft(S, 4) == "Sub ") {
    S = tMid(S, 5);
    isSub = true;
  }
  if (tLeft(S, 9) == "Function ") {
    S = tMid(S, 10);
  }
  AName = RegExNMatch(S, patToken);
  S = Trim(tMid(S, Len(AName) + 1));
  if (tLeft(S, 4) == "Lib ") {
    S = Trim(tMid(S, 5));
    aLib = SplitWord(S, 1);
    S = Trim(tMid(S, Len(aLib) + 1));
    aLib = ReString(aLib);
    if (Left(aLib, 1) == "\"") {
      aLib = Mid(aLib, 2);
    }
    if (Right(aLib, 1) == "\"") {
      aLib = Left(aLib, Len(aLib) - 1);
    }
    if (LCase(Right(aLib, 4)) != ".dll") {
      aLib = aLib + ".dll";
    }
    aLib = LCase(aLib);
  }
  if (tLeft(S, 6) == "Alias ") {
    S = Trim(tMid(S, 7));
    aAlias = SplitWord(S, 1);
    S = Trim(tMid(S, Len(aAlias) + 1));
    aAlias = ReString(aAlias);
    if (Left(aAlias, 1) == "\"") {
      aAlias = Mid(aAlias, 2);
    }
    if (Right(aAlias, 1) == "\"") {
      aAlias = Left(aAlias, Len(aAlias) - 1);
    }
  }
  if (tLeft(S, 1) == "(") {
    S = tMid(S, 2);
  }
  aArgs = nextBy(S, ")");
  S = Trim(tMid(S, Len(aArgs) + 2));
  if (tLeft(S, 3) == "As ") {
    S = Trim(tMid(S, 4));
    aReturn = SplitWord(S, 1);
    S = Trim(tMid(S, Len(aReturn) + 1));
  } else {
    aReturn = "Variant";
  }

  S = "";
  S = S + "[DllImport(\"" + aLib + "\"" + IIf(aAlias == "", "", ", EntryPoint = \"" + aAlias + "\"") + ")] ";
  S = S + IIf(isPrivate, "private ", "public ");
  S = S + "static extern ";
  S = S + IIf(isSub, "void ", ConvertDataType(aReturn)) + " ";
  S = S + AName;
  S = S + "(";
  do {
    if (aArgs == "") {
      break;
    }
    tArg = Trim(nextBy(aArgs, ","));
    aArgs = tMid(aArgs, Len(tArg) + 2);
    S = S + IIf(has, ", ", "") + ConvertParameter(tArg, true);
    has = true;
  } while(!(true));
  S = S + ");";


  ConvertAPIDef = S;
  return ConvertAPIDef;
}

public static string ConvertConstant(string S, bool isGlobal= true) {
  string ConvertConstant = "";
  string cName = "";
  string cType = "";
  string cVal = "";
  bool isPrivate = false;

  if (tLeft(S, 7) == "Public ") {
    S = Mid(Trim(S), 8);
  }
  if (tLeft(S, 7) == "Global ") {
    S = Mid(Trim(S), 8);
  }
  if (tLeft(S, 8) == "Private ") {
    S = Mid(Trim(S), 9);
    isPrivate = true;
  }
  if (tLeft(S, 6) == "Const ") {
    S = Mid(Trim(S), 7);
  }
  cName = SplitWord(S, 1);
  S = Trim(Mid(Trim(S), Len(cName) + 1));
  if (tLeft(S, 3) == "As ") {
    S = Trim(Mid(Trim(S), 3));
    cType = SplitWord(S, 1);
    S = Trim(tMid(S, Len(cType) + 1));
  } else {
    cType = "Variant";
  }

  if (Left(S, 1) == "=") {
    S = Trim(Mid(S, 2));
    cVal() = ConvertValue(S);
  } else {
    cVal() = ConvertDefaultDefault(cType);
  }

  if (cType == "Date") {
    ConvertConstant = IIf(isGlobal, IIf(isPrivate, "private ", "public "), "") + "static readonly " + ConvertDataType(cType) + " " + cName + " = " + cVal() + ";";
  } else {
    ConvertConstant = IIf(isGlobal, IIf(isPrivate, "private ", "public "), "") + "const " + ConvertDataType(cType) + " " + cName + " = " + cVal() + ";";
  }
  return ConvertConstant;
}

public static string ConvertEvent(string S) {
  string ConvertEvent = "";
  string cName = "";
  string cArgs = "";
  string tArgs = "";
  bool isPrivate = false;

  string R = "";
  string N = "";
  string M = "";

  int I = 0;
  int J = 0;

  string A = "";

  if (tLeft(S, 7) == "Public ") {
    S = Mid(Trim(S), 8);
  }
  if (tLeft(S, 8) == "Private ") {
    S = Mid(Trim(S), 9);
    isPrivate = true;
  }
  if (tLeft(S, 6) == "Event ") {
    S = Mid(Trim(S), 7);
  }
  cName = RegExNMatch(S, patToken);
  cArgs = Trim(Mid(Trim(S), Len(cName) + 1));
  if (Left(cArgs, 1) == "(") {
    cArgs = Mid(cArgs, 2);
  }
  if (Right(cArgs, 1) == ")") {
    cArgs = Left(cArgs, Len(cArgs) - 1);
  }

  N = 0;
  do {
    N = N + 1;
    A = nextBy(cArgs, ",", N);
    if (A == "") {
      break;
    }
    tArgs = tArgs + IIf(N == 1, "", ", ");
    tArgs = tArgs + ConvertParameter(A, true);
  } while(!(true));

  N = vbCrLf;
  M = "";
  R = "";
  R = R + M + "public delegate void " + cName + "Handler(" + tArgs + ");";
  R = R + N + "public event " + cName + "Handler event" + cName + ";";

  ConvertEvent = R;
  return ConvertEvent;
}

public static dynamic ConvertEnum(string S) {
  dynamic ConvertEnum = null;
  bool isPrivate = false;
  string EName = "";

  string Res = "";
  bool has = false;

  if (tLeft(S, 7) == "Public ") {
    S = tMid(S, 8);
  }
  if (tLeft(S, 8) == "Private ") {
    S = tMid(S, 9);
    isPrivate = true;
  }
  if (tLeft(S, 5) == "Enum ") {
    S = tMid(S, 6);
  }
  EName = RegExNMatch(S, patToken, 0);
  S = nlTrim(tMid(S, Len(EName) + 1));

  Res = "public enum " + EName + " {";

  while(tLeft(S, 8) != "End Enum" && S != "") {
    EName = RegExNMatch(S, patToken, 0);
    Res = Res + IIf(has, ",", "") + vbCrLf + sSpace(SpIndent) + EName;
    has = true;

    S = nlTrim(tMid(S, Len(EName) + 1));
    if (tLeft(S, 1) == "=") {
      S = nlTrim(Mid(S, 3));
      if (Left(S, 1) == "&") {
        EName = ConvertElement(RegExNMatch(S, "&H[0-9A-F]+"));
      } else {
        EName = RegExNMatch(S, "[0-9]*", 0);
      }
      Res = Res + " = " + EName;
      S = nlTrim(tMid(S, Len(EName) + 1));
    }
  }
  Res = Res + vbCrLf + "}";

  ConvertEnum = Res;
  return ConvertEnum;
}

public static dynamic ConvertType(string S) {
  dynamic ConvertType = null;
  bool isPrivate = false;
  string EName = "";
  string eArr = "";
  string eType = "";

  string Res = "";

  string N = "";

  if (tLeft(S, 7) == "Public ") {
    S = tMid(S, 8);
  }
  if (tLeft(S, 8) == "Private ") {
    S = tMid(S, 9);
    isPrivate = true;
  }
  if (tLeft(S, 5) == "Type ") {
    S = tMid(S, 6);
  }
  EName = RegExNMatch(S, patToken, 0);
  S = nlTrim(tMid(S, Len(EName) + 1));
//If IsInStr(eName, "OSVERSIONINFO") Then Stop

  Res = IIf(isPrivate, "private ", "public ") + "class " + EName + " {";

  while(tLeft(S, 8) != "End Type" && S != "") {
    EName = RegExNMatch(S, patToken, 0);
    S = nlTrim(tMid(S, Len(EName) + 1));
    eArr = "";
    if (LMatch(S, "(")) {
      N = nextBy(Mid(S, 2), ")");
      S = nlTrim(Mid(S, Len(N) + 3));
      N = ConvertValue(N);
      eArr = "[" + N + "]";
    }

    if (tLeft(S, 3) == "As ") {
      S = nlTrim(Mid(S, 4));
      eType = RegExNMatch(S, patToken, 0);
      S = nlTrim(tMid(S, Len(eType) + 1));
    } else {
      eType = "Variant";
    }
    Res = Res + vbCrLf + " public " + ConvertDataType(eType) + IIf(eArr == "", "", "[]") + " " + EName;
    if (eArr == "") {
      Res = Res + " = " + ConvertDefaultDefault(eType);
    } else {
      Res = Res + " = new " + ConvertDataType(eType) + eArr;
    }
    Res = Res + ";";
    if (tLMatch(S, "* ")) {
      S = Mid(LTrim(S), 3);
      N = RegExNMatch(S, "[0-9]+", 0);
      S = nlTrim(Mid(LTrim(S), Len(N) + 1));
      Res = Res + " //TODO: Fixed Length Strings Not Supported: * " + N;
    }

  }
  Res = Res + vbCrLf + "}";

  ConvertType = Res;
  return ConvertType;
}

public static string ConvertParameter(string S, bool NeverUnused= false) {
  string ConvertParameter = "";
  bool isOptional = false;

  bool isByRef = false;
  bool asOut = false;

  string Res = "";

  string pName = "";
  string pType = "";
  string pDef = "";

  string TName = "";


  S = Trim(S);
  if (tLeft(S, 9) == "Optional ") {
    isOptional = true;
    S = Mid(S, 10);
  }
  isByRef = true;
  if (tLeft(S, 6) == "ByVal ") {
    isByRef = false;
    S = Mid(S, 7);
  }
  if (tLeft(S, 6) == "ByRef ") {
    isByRef = true;
    S = Mid(S, 7);
  }
  pName = SplitWord(S, 1);
  if (isByRef && SubParam(pName).AssignedBeforeUsed) {
    asOut = true;
  }
  S = Trim(Mid(S, Len(pName) + 1));
  if (tLeft(S, 2) == "As") {
    S = tMid(S, 4);
    pType = SplitWord(S, 1, "=");
    S = Trim(Mid(S, Len(pType) + 1));
  } else {
    pType = "Variant";
  }
  if (Left(S, 1) == "=") {
    pDef = ConvertValue(Trim(Mid(Trim(S), 2)));
    S = "";
  } else {
    pDef = ConvertDefaultDefault(pType);
  }

  Res = "";
  if (isByRef) {
    Res = Res + IIf(asOut, "out ", "ref ");
  }
  Res = Res + ConvertDataType(pType) + " ";
  if (IsInStr(pName, "()")) {
    Res = Res + "[] ";
    pName = Replace(pName, "()", "");
  }
  TName = pName;
  if (!NeverUnused) {
    if (!SubParam(pName).Used && !(SubParam(pName).Param && SubParam(pName).Assigned)) {
      TName = TName + "_UNUSED";
    }
  }
  Res = Res + TName;
  if (isOptional && !isByRef) {
    Res = Res + "= " + pDef;
  }

  SubParamDecl(pName, pType, false, true, false);
  ConvertParameter = Trim(Res);
  return ConvertParameter;
}

public static string ConvertPrototype(string SS, out string returnVariable, bool asModule= false, out string asName) {
  string ConvertPrototype = "";
  const dynamic retToken = "#RET#";
  string Res = "";

  string fName = "";
  string fArgs = "";
  string retType = "";
  string T = "";

  string tArg = "";

  bool isSub = false;

  bool hArgs = false;

  string S = "";


  S = SS;

  Res = "";
  returnVariable = "";
  isSub = false;
  if (LMatch(S, "Public ")) {
    Res = Res + "public ";
    S = Mid(S, 8);
  }
  if (LMatch(S, "Private ")) {
    Res = Res + "private ";
    S = Mid(S, 9);
  }
  if (LMatch(S, "Friend ")) {
    S = Mid(S, 8);
  }
  if (asModule) {
    Res = Res + "static ";
  }
  if (LMatch(S, "Sub ")) {
    Res = Res + "void ";
    S = Mid(S, 5);
    isSub = true;
  }
  if (LMatch(S, "Function ")) {
    Res = Res + retToken + " ";
    S = Mid(S, 10);
  }

  fName = Trim(SplitWord(Trim(S), 1, "("));
  asName = fName;

  S = Trim(tMid(S, Len(fName) + 2));
  if (Left(S, 1) == "(") {
    S = Trim(tMid(S, 2));
  }
  fArgs = Trim(nextBy(S, ")"));
  S = Mid(S, Len(fArgs) + 2);
  while(Right(fArgs, 1) == "(") {
    fArgs = fArgs + ") ";
    string tMore = "";

    tMore = Trim(nextBy(S, ")"));
    fArgs = fArgs + tMore;
    S = Mid(S, Len(tMore) + 2);
  }
  if (Left(S, 1) == ")") {
    S = Trim(tMid(S, 2));
  }

  if (!isSub) {
    if (tLeft(S, 2) == "As") {
      retType = Trim(Mid(Trim(S), 3));
    } else {
      retType = "Variant";
    }
    if (Right(retType, 1) == ")" && Right(retType, 2) != "()") {
      retType = Left(retType, Len(retType) - 1);
    }
    Res = Replace(Res, retToken, ConvertDataType(retType));
  }

  Res = Res + fName;
  Res = Res + "(";
  hArgs = false;
  do {
    if (Trim(fArgs) == "") {
      break;
    }
    tArg = nextBy(fArgs, ",");
    fArgs = LTrim(Mid(fArgs, Len(tArg) + 2));

    Res = Res + IIf(hArgs, ", ", "");
    if (LMatch(tArg, "ParamArray")) {
      Res = Res + "params ";
      tArg = "ByVal " + Trim(Mid(tArg, 12));
    }
    Res = Res + ConvertParameter(tArg);
    hArgs = true;
  } while(!(Len(fArgs) == 0));

  Res = Res + ") {";
  if (retType != "") {
    returnVariable = fName;
    Res = Res + vbCrLf + sSpace(SpIndent) + ConvertDataType(retType) + " " + returnVariable + " = " + ConvertDefaultDefault(retType) + ";";
    SubParamDecl(returnVariable, retType, false, false, true);
  }

  if (IsEvent(asName)) {
    Res = EventStub(asName) + Res;
  }
  ConvertPrototype = Trim(Res);
  return ConvertPrototype;
}

public static string ConvertCondition(string S) {
  string ConvertCondition = "";
  ConvertCondition = "(" + S + ")";
  return ConvertCondition;
}

public static string ConvertElement(string S) {
  string ConvertElement = "";
//Debug.Print "ConvertElement: " & S
//If IsInStr(S, "frmSetup") Then Stop
//If IsInStr(S, "chkShowBalance.Value") Then Stop
//If IsInStr(S, "optTelephone") Then Stop
  string FirstToken = "";
  string FirstWord = "";

  string T = "";
  bool Complete = false;

  S = Trim(S);
  if (S == "") {
    return ConvertElement;

  }

//If IsInStr(S, "Debug.Print") Then Stop
  if (Left(Trim(S), 2) == "&H") {
    ConvertElement = "0x" + Mid(Trim(S), 3);
    return ConvertElement;

  }

  if (IsNumeric(Trim(S))) {
    ConvertElement = Val(S);
    if (IsInStr(S, ".")) {
      ConvertElement = ConvertElement + "m";
    }
    return ConvertElement;

  }

  int vMax = 0;

  while(RegExTest(S, "#[0-9]+/[0-9]+/[0-9]+#")) {
    string dStr = "";

    dStr = RegExNMatch(S, "#[0-9]+/[0-9]+/[0-9]+#", 0);
    S = Replace(S, dStr, "DateValue(\"" + Mid(dStr, 2, Len(dStr) - 2) + "\")");
    vMax = vMax + 1;
    if (vMax > 10) {
      break;
    }
  }

//If IsInStr(S, "RS!") Then Stop
//If IsInStr(S, ".SetValueDisplay Row") Then Stop
//If IsInStr(S, "cmdSaleTotals.Move") Then Stop
//If IsInStr(S, "2830") Then Stop
//If IsInStr(S, "True") Then Stop
//If IsInStr(S, ":=") Then Stop
//If IsInStr(S, "GetRecordNotFound") Then Stop
//If IsInStr(S, "Nonretro_14day") Then Stop

  S = RegExReplace(S, patNotToken + patToken + "!" + patToken + patNotToken, "$1$2(\"$3\")$4"); // RS!Field -> RS("Field")
  S = RegExReplace(S, "^" + patToken + "!" + patToken + patNotToken, "$1(\"$2\")$3"); // RS!Field -> RS("Field")

  S = RegExReplace(S, "([^a-zA-Z0-9_.])NullDate([^a-zA-Z0-9_.])", "$1NullDate()$2");

  S = ConvertVb6Specific(S, ref Complete);
  if (Complete) {
    ConvertElement = S;
    return ConvertElement;

  }

  if (RegExTest(Trim(S), "^" + patToken + "$")) {
//    If S = "SqFt" Then Stop
    if (IsFuncRef(Trim(S)) && S != CurrSub) {
      ConvertElement = Trim(S) + "()";
      return ConvertElement;

    } else if (IsEnumRef(Trim(S))) {
      ConvertElement = EnumRefRepl(Trim(S));
      return ConvertElement;

    }
  }

//If IsInStr(S, "Not optTagIncoming") Then Stop
  if (IsControlRef(Trim(S), FormName)) {
//If IsInStr(S, "optTagIncoming") Then Stop
    S = FormControlRepl(S, FormName);
  } else if (LMatch(Trim(S), "Not ") && IsControlRef(Mid(Trim(S), 5), FormName)) {
    S = "!(" + FormControlRepl(Mid(Trim(S), 5), FormName) + ")";
  }

  if (IsFormRef(Trim(S))) {
    ConvertElement = FormRefRepl(Trim(S));
    return ConvertElement;

  }


  FirstToken = RegExNMatch(S, patTokenDot, 0);
  FirstWord = SplitWord(S, 1);
  if (FirstWord == "Not") {
    S = "!" + ConvertValue(Mid(S, 5));
    FirstWord = SplitWord(Mid(S, 2));
  }
  if (S == FirstWord) {
    ConvertElement = S;
goto ManageFunctions;
  }
  if (S == FirstToken) {
    ConvertElement = S + "()";
goto ManageFunctions;
  }

  if (FirstToken == FirstWord && !isOperator(SplitWord(S, 2))) { // Sub without parenthesis
    ConvertElement = FirstWord + "(" + SplitWord(S, 2, " ", true, true) + ")";
  } else {
    ConvertElement = S;
  }

ManageFunctions:;
//If IsInStr(ConvertElement, "New_CDbTypeAhead") Then Stop
  if (RegExTest(ConvertElement, "(\\!)?[a-zA-Z0-9_.]+[ ]*\\(.*\\)$")) {
    if ((Left(ConvertElement, 1) == "!")) {
      ConvertElement = "!" + ConvertFunctionCall(Mid(ConvertElement, 2));
    } else {
      ConvertElement = ConvertFunctionCall(ConvertElement);
    }
  }

DoReplacements:;
  if (IsInStr(ConvertElement, ":=")) {
    string Ts = "";

    Ts = SplitWord(ConvertElement, 1, ":=");
    Ts = Ts + ": ";
    Ts = Ts + ConvertElement[SplitWord(ConvertElement, 2, ":=", true, true)];
    ConvertElement = Ts;
  }

  ConvertElement = Replace(ConvertElement, " & ", " + ");
  ConvertElement = Replace(ConvertElement, " = ", " == ");
  ConvertElement = Replace(ConvertElement, "<>", "!=");
  ConvertElement = Replace(ConvertElement, " Not ", " !");
  ConvertElement = Replace(ConvertElement, "(Not ", "(!");
  ConvertElement = Replace(ConvertElement, " Or ", " || ");
  ConvertElement = Replace(ConvertElement, " And ", " && ");
  ConvertElement = Replace(ConvertElement, " Mod ", " % ");
  ConvertElement = Replace(ConvertElement, "Err.", "Err().");
  ConvertElement = Replace(ConvertElement, "Debug.Print", "Console.WriteLine");

  ConvertElement = Replace(ConvertElement, "NullDate", "NullDate");
  while(IsInStr(ConvertElement, ", ,")) {
    ConvertElement = Replace(ConvertElement, ", ,", ", _,");
  }
  ConvertElement = Replace(ConvertElement, "(,", "(_,");

//If IsInStr(ConvertElement, "&H") And Right(ConvertElement, 1) = "&" Then Stop
//If IsInStr(ConvertElement, "1/1/2001") Then Stop

  ConvertElement = RegExReplace(ConvertElement, "([0-9])#", "$1");

  if (Left(ConvertElement, 2) == "&H") {
    ConvertElement = "0x" + Mid(ConvertElement, 3);
    if (Right(ConvertElement, 1) == "&") {
      ConvertElement = Left(ConvertElement, Len(ConvertElement) - 1);
    }
  }

  if (WithLevel > 0) {
    T = Stack(ref WithVars, "##REM##", true);
    ConvertElement = Trim(RegExReplace(ConvertElement, "([ (])(\\.)" + patToken, "$1" + T + "$2$3"));
    if (Left(ConvertElement, 1) == ".") {
      ConvertElement = T + ConvertElement;
    }
  }
  return ConvertElement;
}

public static string ConvertFunctionCall(string fCall) {
  string ConvertFunctionCall = "";
  int I = 0;
  int N = 0;
  string TB = "";
  string Ts = "";
  string TName = "";

  string TV = "";

  Variable vP = null;

//Debug.Print "ConvertFunctionCall: " & fCall

  TB = "";
  TName = RegExNMatch(fCall, "^[a-zA-Z0-9_.]*");
  TB = TB + TName;

  Ts = Mid(fCall, Len(TName) + 2);
  Ts = Left(Ts, Len(Ts) - 1);

  vP = SubParam(TName);
  if (ConvertDataType(vP.asType) == "Recordset") {
    TB = TB + ".Fields[";
    TB = TB + ConvertValue(Ts);
    TB = TB + "].Value";
  } else if (vP.asArray != "") {
    TB = TB + "[";
    TB = TB + ConvertValue(Ts);
    TB = TB + "]";
//    TB = Replace(TB, ", ", "][")
  } else {
    N = nextByPCt(Ts, ",");
    TB = TB + "(";
    for(I=1; I<N; I++) {
      if (I != 1) {
        TB = TB + ", ";
      }
      TV = nextByP(Ts, ",", I);
      if (IsFuncRef(TName)) {
        if (Trim(TV) == "") {
          TB = TB + ConvertElement(FuncRefArgDefault(TName, I));
        } else {
          if (FuncRefArgByRef(TName, I)) {
            TB = TB + "ref ";
          }
          TB = TB + ConvertValue(TV);
        }
      } else {
        TB = TB + ConvertValue(TV);
      }
    }
    TB = TB + ")";
  }
  ConvertFunctionCall = TB;
  return ConvertFunctionCall;
}

public static string ConvertValue(string S) {
  string ConvertValue = "";
  string F = "";
  string Op = "";
  string OpN = "";

  string O = "";

  O = "";
  S = Trim(S);
  if (S == "") {
    return ConvertValue;

  }

//If IsInStr(S, "GetMaxFieldValue") Then Stop
//If IsInStr(S, "DBAccessGeneral") Then Stop
//If IsInStr(S, "tallable") Then Stop
//If Left(S, 3) = "RS(" Then Stop
//If Left(S, 6) = "DBName" Then Stop
//If Left(S, 6) = "fName" Then Stop

  SubParamUsedList(TokenList(S));

  if (RegExTest(S, "^-[a-zA-Z0-9_]")) {
    ConvertValue = "-" + ConvertValue[Mid(S, 2)];
    return ConvertValue;

  }

  while(true) {
    F = NextByOp(S, 1, ref Op);
    if (F == "") {
      break;
    }
    switch(Trim(Op)) {
      case "\\":
        OpN = "/";
        break;
      case "=":
        OpN = " == ";
        break;
      case "<>":
        OpN = " != ";
        break;
      case "&":
        OpN = " + ";
        break;
      case "Mod":
        OpN = " % ";
        break;
      case "Is":
        OpN = " == ";
        break;
      case "Like":
        OpN = " == ";
        break;
      case "And":
        OpN = " && ";
        break;
      case "Or":
        OpN = " || ";
        break;
      default:
        OpN = Op;
break;
}


    if (Left(F, 1) == "("& Right(F, 1) == ")") {
      O = O + "(" + ConvertValue[Mid(F, 2, Len(F) - 2)] + ")" + OpN;
    } else {
      O = O + ConvertElement(F) + OpN;
    }

    if (Op == "") {
      break;
    }
    S = Mid(S, Len(F) + Len(Op) + 1);
    if (S == "" || Op == "") {
      break;
    }
  }
  ConvertValue = O;
  return ConvertValue;
}

public static string ConvertGlobals(string Str, bool asModule= false) {
  string ConvertGlobals = "";
  string Res = "";

  dynamic S = null;
  dynamic L = null;
  string O = "";

  int Ind = 0;

  string Building = "";

  int inCase = 0;

  string returnVariable = "";

  int N = 0;


  Res = "";
  Building = "";
  Str = Replace(Str, vbLf, "");
  S = Split(Str, vbCr);
  Ind = 0;
  N = 0;
//  Prg 0, UBound(S) - LBound(S) + 1, "Globals..."
  InitDeString();
  foreach(var iterL in S) {
L = iterL;
    L = DeComment(L);
    L = DeString(L);
    O = "";
    if (Building != "") {
      Building = Building + vbCrLf + L;
      if (tLeft(L, 8) == "End Type") {
        O = ConvertType(Building);
        Building = "";
      } else if (tLeft(L, 8) == "End Enum") {
        O = ConvertEnum(Building);
        Building = "";
      }
    } else if (L == "Option *") {
      O = "// " + L;
    } else if (RegExTest(L, "^(Public |Private |)Declare ")) {
      O = ConvertAPIDef(L);
    } else if (RegExTest(L, "^(Global |Public |Private |)Const ")) {
      O = ConvertConstant(L, true);
    } else if (RegExTest(L, "^(Public |Private |)Event ")) {
      O = ConvertEvent(L);
    } else if (RegExTest(L, "^(Public |Private |)Enum ")) {
      Building = L;
    } else if (RegExTest(LTrim(L), "^(Public |Private |)Type ")) {
      Building = L;
    } else if (tLeft(L, 8) == "Private " || tLeft(L, 7) == "Public " || tLeft(L, 4) == "Dim ") {
      O = ConvertDeclare(L, 0, true, asModule);
    }

    O = ReComment(O);
    Res = Res + ReComment(O) + IIf(O == "" || Right(O, 2) == vbCrLf, "", vbCrLf);
    N = N + 1;
//    Prg N
//    If N Mod 10000 = 0 Then Stop
  }
//  Prg

  Res = ReString(Res, true);
  ConvertGlobals = Res;
  return ConvertGlobals;
}

public static string ConvertCodeLine(string S) {
  string ConvertCodeLine = "";
  int T = 0;
  string A = "";
  string B = "";

  string P = "";
  Variable V = null;

  string FirstWord = "";
  string Rest = "";

  int N = 0;


//If IsInStr(S, "dbClose") Then Stop
//If IsInStr(S, "Nothing") Then Stop
//If IsInStr(S, "Close ") Then Stop
//If IsInStr(S, "& functionType & fieldInfo &") Then Stop
//If IsInStr(S, " & vbCrLf2 & Res)") Then Stop
//If IsInStr(S, "Res = CompareSI(SI1, SI2)") Then Stop
//If IsInStr(S, "frmPrintPreviewDocument") Then Stop
//If IsInStr(S, "NewAudit.Name1") Then Stop
//If IsInStr(S, "optDelivered") Then Stop
//If IsInStr(S, " Is Nothing Then") Then Stop
//If IsInStr(S, "SqFt, SqYd") Then Stop
//If IsInStr(S, "optTagIncoming") Then Stop
//If IsInStr(S, "Kill modAshleyItemAlign") Then Stop
//If IsInStr(S, "PRFolder") Then Stop

  if (Trim(S) == "") {
    ConvertCodeLine = "";
    return ConvertCodeLine;

  }
  bool Complete = false;

  S = ConvertVb6Specific(S, ref Complete);
  if (Complete) {
    ConvertCodeLine = S;
    return ConvertCodeLine;

  }

  if (RegExTest(Trim(S), "^[a-zA-Z0-9_.()]+ \\= ") || RegExTest(Trim(S), "^Set [a-zA-Z0-9_.()]+ \\= ")) {
    T = InStr(S, "=");
    A = Trim(Left(S, T - 1));
    if (tLeft(A, 4) == "Set ") {
      A = Trim(Mid(A, 5));
    }
    SubParamAssign(RegExNMatch(A, patToken));
    if (RegExTest(A, "^" + patToken + "\\(\"[^\"]+\"\\)")) {
      P() = RegExNMatch(A, "^" + patToken);
      V = SubParam(P());
      if (V.Name == P()) {
        SubParamAssign(P());
        switch(V.asType) {
          case "Recordset":
            ConvertCodeLine = RegExReplace(A, "^" + patToken + "(\\(\")([^\"]+)(\"\\))", "$1.Fields[\"$3\"].Value");
            break;
          default:
            if (Left(A, 1) == ".") {
              A = Stack(ref WithVars, "##REM##", true) + A;
            }
            ConvertCodeLine = A;
break;
}
      }
    } else {
      if (Left(A, 1) == ".") {
        A = Stack(ref WithVars, "##REM##", true) + A;
      }
      ConvertCodeLine = A;
    }

    string tAWord = "";

    tAWord = SplitWord(A, 1, ".");
    if (IsFormRef(tAWord)) {
      A = Replace(A, tAWord, tAWord + ".instance", _, 1);
    }

    ConvertCodeLine = ConvertValue(ConvertCodeLine) + " = ";

    B = ConvertValue(Trim(Mid(S, T + 1)));
    ConvertCodeLine = ConvertCodeLine + B;
  } else {
//Debug.Print S
    FirstWord = SplitWord(Trim(S));
    Rest = SplitWord(Trim(S), 2, " ", true, true);
    if (Rest == "") {
      ConvertCodeLine = S + IIf(Right(S, 1) != ")", "()", "");
      ConvertCodeLine = ConvertElement(ConvertCodeLine);
    } else if (FirstWord == "RaiseEvent") {
      ConvertCodeLine = ConvertValue(S);
    } else if (FirstWord == "Debug.Print") {
      ConvertCodeLine = "Console.WriteLine(" + ConvertValue(Rest) + ")";
    } else if (StrQCnt(FirstWord, "(") == 0) {
      ConvertCodeLine = "";
      ConvertCodeLine = ConvertCodeLine + FirstWord + "(";
      N = 0;
      do {
        N = N + 1;
        B = nextByP(Rest, ", ", N);
        if (B == "") {
          break;
        }
        ConvertCodeLine = ConvertCodeLine + IIf(N == 1, "", ", ") + ConvertValue(B);
      } while(!(true));
      ConvertCodeLine = ConvertCodeLine + ")";
//      ConvertCodeLine = ConvertElement(ConvertCodeLine)
    } else {
      ConvertCodeLine = ConvertValue(S);
    }
    if (WithLevel > 0& Left(Trim(ConvertCodeLine), 1) == ".") {
      ConvertCodeLine = Stack(ref WithVars, "##REM##", true) + Trim(ConvertCodeLine);
    }
  }

//  If IsInStr(ConvertCodeLine, ",,,,,,,") Then Stop

  ConvertCodeLine = ConvertCodeLine + ";";
//Debug.Print ConvertCodeLine
  return ConvertCodeLine;
}

public static string PostConvertCodeLine(string Str) {
  string PostConvertCodeLine = "";
  string S = "";

  S = Str;

//  If IsInStr(S, "optPoNo") Then Stop

  if (IsInStr(S, "0 &")) {
    S = Replace(S, "0 &", "0");
  }
  if (IsInStr(S, ".instance.instance")) {
    S = Replace(S, ".instance.instance", ".instance");
  }
  if (IsInStr(S, ".IsChecked)")) {
    S = Replace(S, ".IsChecked)", ".IsChecked == true)", 1);
  }
  if (IsInStr(S, ".IsChecked &")) {
    S = Replace(S, ".IsChecked", ".IsChecked == true", 1);
  }
  if (IsInStr(S, ".IsChecked |")) {
    S = Replace(S, ".IsChecked", ".IsChecked == true", 1);
  }
  if (IsInStr(S, ".IsChecked,")) {
    S = Replace(S, ".IsChecked", ".IsChecked == true", 1);
  }
  if (IsInStr(S, ".IsChecked == 1,")) {
    S = Replace(S, ".IsChecked == 1", ".IsChecked == true", 1);
  }
  if (IsInStr(S, ".IsChecked == 0,")) {
    S = Replace(S, ".IsChecked == 1", ".IsChecked == false", 1);
  }

  if (IsInStr(S, ".Visibility = true")) {
    S = Replace(S, ".Visibility = true", ".setVisible(true)");
  }
  if (IsInStr(S, ".Visibility = false")) {
    S = Replace(S, ".Visibility = false", ".setVisible(false)");
  }

  if (IsInStr(S, ".Print(")) {
    if (IsInStr(S, ";);")) {
      S = Replace(S, ";);", ");");
      S = Replace(S, "Print(", "PrintNNL(");
    }
    S = Replace(S, "; ", ", ");
  }
  if (IsInStr(S, ".Line((")) {
    S = Replace(S, ") - (", ", ");
    S = Replace(S, "Line((", "Line(");
    S = Replace(S, "));", ");");
  }

  S = Replace(S, "vbRetryCancel +", "vbRetryCancel |");
  S = Replace(S, "vbOkOnly +", "vbOkOnly |");
  S = Replace(S, "vbOkCancel +", "vbOkCancel |");
  S = Replace(S, "vbExclamation +", "vbExclamation |");
  S = Replace(S, "vbYesNo +", "vbYesNo |");
  S = Replace(S, "vbQuestion +", "vbQuestion |");
  S = Replace(S, "vbOKCancel +", "vbOKCancel |");
  S = Replace(S, "+ vbExclamation", "| vbExclamation");

  PostConvertCodeLine = S;
  return PostConvertCodeLine;
}
}
