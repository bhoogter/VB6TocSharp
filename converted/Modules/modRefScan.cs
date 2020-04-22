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
using static VB2CS.Forms.frm;
using static VB2CS.Forms.frmConfig;


static class modRefScan {
// Option Explicit
private static string OutRes = "";
private static string cFuncRef_Name = "";
private static string cFuncRef_Value = "";
private static string cEnuRef_Name = "";
private static string cEnumRef_Value = "";
private static Collection Funcs = null;
private static Collection LocalFuncs = null;


private static string RefList(bool KillRef= false) {
  string RefList = "";
  // TODO (not supported): On Error Resume Next
  RefList = App.Path + "\\refs.txt";
  if (KillRef) {
    File.Delete(RefList);();
  }
  return RefList;
}

public static int FuncsCount(bool vLocal= false) {
  int FuncsCount = 0;
  // TODO (not supported): On Error Resume Next
  if (vLocal) {
    FuncsCount = LocalFuncs.Count;
  } else {
    FuncsCount = Funcs.Count;
  }
  return FuncsCount;
}

public static int ScanRefs() {
  int ScanRefs = 0;
  dynamic L = null;
  string T = "";

  // TODO (not supported): On Error Resume Next
  OutRes = "";
  ScanRefs = 0;
  foreach(var L in Split(VBPModules(vbpFile), vbCrLf)) {
    if (L == "") {
goto SkipMod;
    }
    ScanRefs = ScanRefs + ScanRefsFile(FilePath(vbpFile) + L);
SkipMod:
  }

  foreach(var L in Split(VBPForms(vbpFile), vbCrLf)) {
    L = Replace(L, ".frm", "");
    if (L == "") {
goto SkipForm;
    }
    T = vbCrLf + L + ":" + L + ":Form:";
    OutRes = OutRes + T;

//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    string S = "";
    int J = 0;
    string Preamble = "";
    string ControlRefs = "";

    S = ReadEntireFile(vbpPath + L + ".frm");
    J = CodeSectionLoc(S);
    Preamble = Left(S, J - 1);
    ControlRefs = FormControls(L, Preamble, false);
    OutRes = OutRes + ControlRefs;
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ScanRefs = ScanRefs + 1;
SkipForm:
  }
  RefList(KillRef: true);
  WriteFile(RefList, OutRes);
  OutRes = "";
  return ScanRefs;
}

private static int ScanRefsFile(string FN) {
  int ScanRefsFile = 0;
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
  ScanRefsFile = 0;
  foreach(var LL in Split(S, vbCrLf)) {
    DoCont = Right(LL, 1) == "_";
    if (!Cont && !DoCont) {
      L = Trim(LL);
      Cont = false;
    } else if (Cont && !DoCont) {
      L = L + Trim(LL);
      Cont = false;
    } else if (!Cont && DoCont) {
      L = Trim(Left(LL, Len(LL) - 2));
      Cont = true;
goto ;
    } else if (Cont && DoCont) {
      L = L + Trim(Left(LL, Len(LL) - 2));
      Cont = true;
goto ;
    }

    if (tLMatch(L, "Function ") || tLMatch(L, "Public Function ") || tLMatch(L, "Sub ") || tLMatch(L, "Public Sub ") || false) {
      F = Trim(L);
      if (Left(F, 7) == "Public ") {
        F = Mid(F, 8);
      }
      F = Trim(nextBy(F, ":"));

      G = F;
      if (tLMatch(G, "Function ")) {
        G = Mid(G, 10);
      }
      if (tLMatch(G, "Sub ")) {
        G = Mid(G, 5);
      }
      G = nextBy(G, "(");

      F = M + ":" + G + ":Function:" + F;
      OutRes = OutRes + vbCrLf + F;
      ScanRefsFile = ScanRefsFile + 1;
    } else if (tLMatch(L, "Declare ") || tLMatch(L, "Public Decalre ")) {
      L = LTrim(L);
      if (LMatch(L, "Public ")) {
        L = Mid(L, 8);
      }
      if (LMatch(L, "Declare ")) {
        L = Mid(L, 9);
      }
      G = SplitWord(L);

    } else if (tLMatch(L, "Const ") || tLMatch(L, "Public Const ") || tLMatch(L, "Global Const ")) {
      L = LTrim(L);
      if (LMatch(L, "Public ")) {
        L = Mid(L, 8);
      }
      if (LMatch(L, "Global ")) {
        L = Mid(L, 8);
      }
      if (LMatch(L, "Const ")) {
        L = Mid(L, 7);
      }
      G = SplitWord(L);
    } else if (tLMatch(L, "Enum ") || tLMatch(L, "Public Enum ")) {
      L = LTrim(L);
      if (LMatch(L, "Public ")) {
        L = Mid(L, 8);
      }
      if (LMatch(L, "Enum ")) {
        L = Mid(L, 5);
      }
      CurrEnum = Trim(L);
    } else if (tLMatch(L, "End Enum")) {
      CurrEnum = "";
    } else if (CurrEnum != "") {
      G = SplitWord(L);
      F = M + ":" + G + ":Enum:" + CurrEnum + "." + G;
      OutRes = OutRes + vbCrLf + F;
      ScanRefsFile = ScanRefsFile + 1;
    }
NextLine:
  }
  return ScanRefsFile;
}

public static string ScanRefsFileToString(string FN) {
  string ScanRefsFileToString = "";
  OutRes = "";
  ScanRefsFile(FN);
  ScanRefsFileToString = OutRes;
  OutRes = "";
  return ScanRefsFileToString;
}

private static void InitFuncs() {
  string S = "";
  dynamic L = null;

  if (Dir(RefList) == "") {
    ScanRefs();
  }
  if (!(Funcs Is Nothing)) {
return;

  }
  S = ReadEntireFile(RefList);
  Funcs = new Collection();;
  // TODO (not supported): On Error Resume Next
  foreach(var L in Split(S, vbCrLf)) {
    Funcs.Add(L, SplitWord(L, 2, ":"));
  }
  InitLocalFuncs();
}

public static void InitLocalFuncs(string S_UNUSED= "") {
  // TODO (not supported): On Error Resume Next
  dynamic L = null;

  LocalFuncs = new Collection();;
  foreach(var L in Split(S, vbCrLf)) {
    LocalFuncs.Add(L, SplitWord(L, 2, ":"));
  }
}

public static string FuncRef(string FName) {
  string FuncRef = "";
  if (FName == cFuncRef_Name) {
    FuncRef = cFuncRef_Value;
    return FuncRef;

  }

  InitFuncs();
  // TODO (not supported): On Error Resume Next
  FuncRef = Funcs(FName);
  if (FuncRef == "") {
    FuncRef = LocalFuncs(FName);
  }

  cFuncRef_Name = FName;
  cFuncRef_Value = FuncRef;
  return FuncRef;
}

public static string FuncRefModule(string FName) {
  string FuncRefModule = "";
  FuncRefModule = nextBy(FuncRef(FName), ":");
  return FuncRefModule;
}

public static string FuncRefEntity(string FName) {
  string FuncRefEntity = "";
  FuncRefEntity = nextBy(FuncRef(FName), ":", 3);
  return FuncRefEntity;
}

public static string FuncRefDecl(string FName) {
  string FuncRefDecl = "";
  FuncRefDecl = nextBy(FuncRef(FName), ":", 4);
  return FuncRefDecl;
}

public static bool IsFuncRef(string FName) {
  bool IsFuncRef = false;
  IsFuncRef = FuncRef(FName) != "" && FuncRefEntity(FName) == "Function";
  return IsFuncRef;
}

public static bool IsEnumRef(string FName) {
  bool IsEnumRef = false;
  IsEnumRef = FuncRef(FName) != "" && FuncRefEntity(FName) == "Enum";
  return IsEnumRef;
}

public static bool IsFormRef(string FName) {
  bool IsFormRef = false;
  string T = "";

  T = SplitWord(FName, 1, ".");
  IsFormRef = FuncRef(T) != "" && FuncRefEntity(T) == "Form";
  return IsFormRef;
}

public static bool IsControlRef(string Src, string FormName= "") {
  bool IsControlRef = false;
  string Tok = "";
  string Tok2 = "";

  string FTok = "";
  string TTok = "";

  Tok = RegExNMatch(Src, patToken);
  Tok2 = RegExNMatch(Src, patToken, 1);
  TTok = Tok + "." + Tok2;
  FTok = FormName + "." + Tok;
//If IsInStr(Src, "SetFocus") Then Stop
  if (FuncRef(TTok) != "" && FuncRefEntity(TTok) == "Control" || FuncRef(FTok) != "" && FuncRefEntity(FTok) == "Control") {
    IsControlRef = true;
  }
  return IsControlRef;
}

public static string FuncRefDeclTyp(string FName) {
  string FuncRefDeclTyp = "";
  FuncRefDeclTyp = SplitWord(FuncRefDecl(FName), 1);
  return FuncRefDeclTyp;
}

public static string FuncRefDeclRet(string FName) {
  string FuncRefDeclRet = "";
  FuncRefDeclRet = FuncRefDecl(FName);
  FuncRefDeclRet = Trim(Mid(FuncRefDeclRet, InStrRev(FuncRefDeclRet, " ")));
  if (Right(FuncRefDeclRet, 1) == ")" && Right(FuncRefDeclRet, 2) != "()") {
    FuncRefDeclRet = "";
  }
  return FuncRefDeclRet;
}

public static string FuncRefDeclArgs(string FName) {
  string FuncRefDeclArgs = "";
  // TODO (not supported): On Error Resume Next
  FuncRefDeclArgs = FuncRefDecl(FName);
  FuncRefDeclArgs = Mid(FuncRefDeclArgs, InStr(FuncRefDeclArgs, "(") + 1);
  FuncRefDeclArgs = Left(FuncRefDeclArgs, InStrRev(FuncRefDeclArgs, ")") - 1);
  FuncRefDeclArgs = Trim(FuncRefDeclArgs);
  return FuncRefDeclArgs;
}

public static string FuncRefDeclArgN(string FName, int N) {
  string FuncRefDeclArgN = "";
  string F = "";

  F = FuncRefDeclArgs(FName);
  FuncRefDeclArgN = nextBy(F, ", ", N);
  return FuncRefDeclArgN;
}

public static int FuncRefDeclArgCnt(string FName) {
  int FuncRefDeclArgCnt = 0;
  string F = "";
  string K = "";

  F = FuncRefDeclArgs(FName);
  FuncRefDeclArgCnt = 0;
  do {
    K = nextBy(F, ", ", FuncRefDeclArgCnt + 1);
    if (K == "") {
      return FuncRefDeclArgCnt;

    }
    FuncRefDeclArgCnt = FuncRefDeclArgCnt + 1;
  } while(!(true);
  return FuncRefDeclArgCnt;
}

public static string FuncRefArgType(string FName, int N) {
  string FuncRefArgType = "";
  FuncRefArgType = FuncRefDeclArgN(FName, N);
  if (FuncRefArgType == "") {
    return FuncRefArgType;

  }
  FuncRefArgType = SplitWord(FuncRefArgType, 2, " As ");
  return FuncRefArgType;
}

public static bool FuncRefArgByRef(string FName, int N) {
  bool FuncRefArgByRef = false;
  FuncRefArgByRef = !IsInStr(FuncRefDeclArgN(FName, N), "ByVal ");
  return FuncRefArgByRef;
}

public static bool FuncRefArgOptional(string FName, int N) {
  bool FuncRefArgOptional = false;
  FuncRefArgOptional = IsInStr(FuncRefDeclArgN(FName, N), "Optional ");
  return FuncRefArgOptional;
}

public static string FuncRefArgDefault(string FName, int N) {
  string FuncRefArgDefault = "";
  string aTyp = "";

  if (!FuncRefArgOptional(FName, N)) {
    return FuncRefArgDefault;

  }
  FuncRefArgDefault = SplitWord(FuncRefDeclArgN(FName, N), 2, " = ", true, true);
  if (FuncRefArgDefault == "") {
    FuncRefArgDefault = ConvertDefaultDefault(FuncRefArgType(FName, N));
  }
  return FuncRefArgDefault;
}

public static string EnumRefRepl(string EName) {
  string EnumRefRepl = "";
  EnumRefRepl = FuncRefDecl(EName);
  return EnumRefRepl;
}

public static string FormRefRepl(string FName) {
  string FormRefRepl = "";
  string T = "";
  string U = "";

  T = SplitWord(FName, 1, ".");
  U = FuncRefModule(T) + ".instance";
  FormRefRepl = Replace(FName, T, U);
  return FormRefRepl;
}

public static string FormControlRepl(string Src, string FormName= "") {
  string FormControlRepl = "";
  string Tok = "";
  string Tok2 = "";
  string Tok3 = "";

  string F = "";
  string V = "";

  Tok = RegExNMatch(Src, patToken);
  Tok2 = RegExNMatch(Src, patToken, 1);
  Tok3 = RegExNMatch(Src, patToken, 2);

//If IsInStr(Tok, "BillOSale") Then Stop
//If IsInStr(Src, "SetFocus") Then Stop

  if (!IsFormRef(Tok)) {
    F = Tok;
    V = ConvertControlProperty(F, Tok2, FuncRefDecl(FormName + "." + Tok));
    if (Tok2 != "") {
      FormControlRepl = Replace(Src, Tok2, V);
    } else {
      FormControlRepl = Src + "." + V;
    }
  } else {
    F = Tok + "." + Tok2;
    V = ConvertControlProperty(F, Tok3, FuncRefDecl(Tok + "." + Tok2));
    if (Tok3 != "") {
      FormControlRepl = Replace(Src, Tok3, V);
    } else {
      FormControlRepl = Src + "." + V;
    }
  }
  return FormControlRepl;
}
}
