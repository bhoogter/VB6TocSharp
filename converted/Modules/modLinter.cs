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


static class modLinter {
// Option Explicit
//:::: modLinter
//:::SUMMARY
//: Lint VB6 files
//:
//:::DESCRIPTION
//: Inspect VB6 files to linting.  I.e., format, spacing, and other code-quality features.
//:
//::: Rules
//: - Indentation
//: - File Names
//:
//:::SEE ALSO
//:    - modXML
public static bool LintForBuild = false;
private const int lintFileShort_Len = 20;
private const int lintLint_MaxErrors = 30;
private const int lintLint_TabWidth = 2;
private const string lintLint_IndentContextDiv = ":";
private const int lintLint_MaxBlankLines = 5;
private const int lintLint_MaxBlankLines_AtClose = 2;
private const int lintNoLint_ScanRange = 10;
private const int lintTag_ScanRange = 10;
private const string lintTag_Key = "@";
private const string lintTag_Start = "'" + lintTag_Key;
private const string lintTag_Div = "-";
private const string lintTag_NoLint = lintTag_Start + "NO-LINT";
private const string lintFile_Option = "Option ";
private const string lintFile_Option_Explicit = "Explicit";
private const int lintDotsPerRow = 60;
private const string lintFixList_Sep = "/||\\";
private const string lintFixList_Div = ":oo:";
public enum lintErrorTypes {
  ltUnkn,
  ltLErr,
  ltIdnt,
  ltDECL,
  ltDEPR,
  ltSTOP,
  ltWITH,
  ltVarN,
  ltArgN,
  ltFunN,
  ltCtlN,
  ltObjN,
  ltSelf,
  ltType,
  ltWhtS,
  ltOptn,
  ltBadC,
  ltNTyp,
  ltNOpD
}


private static bool CheckNoLint(string FileName, lintErrorTypes lType= lintErrorTypes.ltUnkn, string vLine= "") {
  bool CheckNoLint = false;
  int I = 0;
  string L = "";
  int A = 0;

  string CA = "";
  string cP = "";
  string cB = "";

  CheckNoLint = false;
  if (LintAbbr(lType) == "") {
    CheckNoLint = true; // Disable lint type
    return CheckNoLint;

  }

  CA = lintTag_NoLint;
  cP = lintTag_NoLint + lintTag_Div;
  cB = lintTag_NoLint + lintTag_Div + UCase(LintAbbr(lType));

  if (IsInStr(vLine, Mid(CA, 2)) && !IsInStr(vLine, Mid(cP, 2))) {
    CheckNoLint = true;
    return CheckNoLint;

  }
  if (IsInStr(vLine, Mid(cB, 2))) {
    CheckNoLint = true;
    return CheckNoLint;

  }

  A = LintModuleFirstLine(FileName());
  for(I=A; I<A + lintTag_ScanRange; I++) {
    L = UCase(ReadFile(FileName(), I, 1));
    if (lType == lintErrorTypes.ltUnkn) {
      if (LMatch(L, CA) && !LMatch(L, cP)) {
        CheckNoLint = true;
        return CheckNoLint;

      }
    } else {
      if (LMatch(L, cB)) {
        CheckNoLint = true;
        return CheckNoLint;

      }
    }
  }
  return CheckNoLint;
}

private static string LintAbbr(lintErrorTypes lType, out string TypeName) {
  string LintAbbr = "";
// if this function returns "", the lint type is ignored.  Add the following after the normal lint type to disable:
// : LintAbbr = ""
  switch(lType) {
    case lintErrorTypes.ltLErr:
      LintAbbr = "LErr";
      TypeName = "Lint Runtime Error";
      break;
    case lintErrorTypes.ltIdnt:
      LintAbbr = "Idnt";
      TypeName = "Indent";
      break;
    case lintErrorTypes.ltDECL:
      LintAbbr = "Decl";
      TypeName = "Declaration";
      break;
    case lintErrorTypes.ltDEPR:
      LintAbbr = "Depr";
      TypeName = "Deprecated";
      break;
    case lintErrorTypes.ltSTOP:
      LintAbbr = "STOP";
      TypeName = "Stop Encountered";
      break;//: LintAbbr = IIf(Not LintForBuild, LintAbbr, "")
    case lintErrorTypes.ltWITH:
      LintAbbr = "WITH";
      TypeName = "With Statement";
      break;
    case lintErrorTypes.ltVarN:
      LintAbbr = "VarN";
      TypeName = "Variable Name";
      break;
    case lintErrorTypes.ltArgN:
      LintAbbr = "ArgN";
      TypeName = "Argument Name";
      break;
    case lintErrorTypes.ltFunN:
      LintAbbr = "FunN";
      TypeName = "Function Name";
      break;
    case lintErrorTypes.ltCtlN:
      LintAbbr = "CtlN";
      TypeName = "Control Name";
      break;
    case lintErrorTypes.ltObjN:
      LintAbbr = "ObjN";
      TypeName = "Object Name";
      break;
    case lintErrorTypes.ltSelf:
      LintAbbr = "Self";
      TypeName = "Self Reference";
      break;
    case lintErrorTypes.ltType:
      LintAbbr = "Type";
      TypeName = "Data Type";
      break;
    case lintErrorTypes.ltWhtS:
      LintAbbr = "WhtS";
      TypeName = "White Space";
      break;
    case lintErrorTypes.ltOptn:
      LintAbbr = "Optn";
      TypeName = "Option";
      break;
    case lintErrorTypes.ltBadC:
      LintAbbr = "BadC";
      TypeName = "Bad Code";

      break;//: LintAbbr = IIf(Not LintForBuild, LintAbbr, "")
    case lintErrorTypes.ltNTyp:
      LintAbbr = "NTyp";
      TypeName = "No Type";
      break;
    case lintErrorTypes.ltNOpD:
      LintAbbr = "NOpD";
      TypeName = "No Default For Optional";
      LintAbbr = "";

      break;
    default:
      LintAbbr = "UNKN";
      TypeName = "Unknown";
break;
}
  return LintAbbr;
}

private static string LintName(lintErrorTypes lType) {
  string LintName = "";
  LintAbbr(lType, LintName);
  return LintName;
}

private static string LintFileShort(string FFile) {
  string LintFileShort = "";
  LintFileShort = AlignString(FileName(FFile), lintFileShort_Len);
  return LintFileShort;
}

private static string AddErrStr(ref string ErrStr, string FileName, string LineNo, string vLine, string Msg, lintErrorTypes lType) {
  string AddErrStr = "";
  Static(ErrCnt(As(Long)));
  if (CheckNoLint(FileName(), lType, vLine)) {
    return AddErrStr;

  }

  if (ErrStr == "") {
    ErrCnt = 0;
  }
  ErrCnt = ErrCnt + 1;
  if (ErrCnt > lintLint_MaxErrors) {
    if (Right(ErrStr, 4) != " >>>") {
      ErrStr = ErrStr + vbCrLf + "<<< Max Error Count Exceeded >>>";
    }
    return AddErrStr;

  }
  if (ErrStr != "") {
    ErrStr = ErrStr + vbCrLf;
  }
  ErrStr = ErrStr + LintFileShort(FileName()) + " (Line " + LineNo + "): " + LintAbbr(lType) + " - " + Msg;
  return AddErrStr;
}

private static bool AddIndent(out int Lvl, out string Str, out string Context, bool POP= false) {
  bool AddIndent = false;
  AddIndent = true;
  Context = Replace(Context, lintLint_IndentContextDiv, "-");
  if (POP) {
    Lvl = Lvl - lintLint_TabWidth;
    if (Lvl < 0) {
      Lvl = 0;
      Str = "";
      AddIndent = false;
    }
    Context = SplitWord(Str, -1, lintLint_IndentContextDiv);
    Str = Left(Str, Len(Str) - Len(Context));
    if (Right(Str, Len(lintLint_IndentContextDiv)) == lintLint_IndentContextDiv) {
      Str = Left(Str, Len(Str) - Len(lintLint_IndentContextDiv));
    }
  } else {
    Lvl = Lvl + lintLint_TabWidth;
    Str = Str + IIf(Str == "", "", lintLint_IndentContextDiv) + Context;
  }
  return AddIndent;
}

private static string IndentContext(string Str) {
  string IndentContext = "";
  IndentContext = SplitWord(Str, -1, lintLint_IndentContextDiv);
  return IndentContext;
}

private static string DeComment(string S) {
  string DeComment = "";
  int I = 0;

  string C = "";

  bool Q = false;

  Q = false;
  DeComment = S;
  if (IsNotInStr(S, "'")) {
    return DeComment;

  }

  for(I=1; I<Len(S); I++) {
    C = Mid(S, I, 1);
    if (C == "\"") {
      Q = !Q;
    }
    if (!Q && C == "'") {
      DeComment = RTrim(Left(S, I - 1));
      return DeComment;

    }
  }
  return DeComment;
}

private static string DeString(string S) {
  string DeString = "";
  const string Q = "\"";
  const string Token = "_";
  int A = 0;
  int B = 0;

  DeString = S;
  A = InStr(S, Q);
  if (A > 0) {
    B = InStr(A + 1, S, Q);
    if (B > 0) {
      DeString = DeString[Left(S, A - 1) + Token + Mid(S, B + 1)];
      return DeString;

    }
  }
  DeString = S;
  return DeString;
}

private static string DeSpace(string S) {
  string DeSpace = "";
  int N = 0;

  DeSpace = S;
  do {
    N = Len(DeSpace);
    DeSpace = Replace(DeSpace, "  ", " ");
    if (Len(DeSpace) == N) {
      return DeSpace;

    }
  } while(!(true);
  return DeSpace;
}

public static bool LintFolder(string Folder_UNUSED= "", bool AutoFix= false, bool ForBuild_UNUSED= false) {
  bool LintFolder = false;
  LintForBuild = true;
  LintFolder = LintFileList(VBPModules(vbpFile) + vbCrLf + VBPClasses(vbpFile) + vbCrLf + VBPForms(), AutoFix);
  return LintFolder;
}

public static bool LintModules(string Folder_UNUSED= "", bool AutoFix= false) {
  bool LintModules = false;
  LintModules = LintFileList(VBPModules(), AutoFix);
  return LintModules;
}

public static bool LintClasses(string Folder_UNUSED= "", bool AutoFix= false) {
  bool LintClasses = false;
  LintClasses = LintFileList(VBPClasses(), AutoFix);
  return LintClasses;
}

public static bool LintForms(string Folder_UNUSED= "", bool AutoFix= false) {
  bool LintForms = false;
  LintForms = LintFileList(VBPForms(), AutoFix);
  return LintForms;
}

public static bool LintFileList(string List_UNUSED, bool AutoFix) {
  bool LintFileList = false;
  string E = "";
  dynamic L = null;

  int X = 0;

  DateTime StartTime = DateTime.MinValue;

  StartTime = DateTime.Now;;

  foreach(var L in Split(List, vbCrLf)) {
    if (!LintFile(L, ref E)) {
      if (AutoFix) {
        LintFileIndent(DevelopmentFolder() + L);
        Debug.PrintNNL("x");
      } else {
        Debug.Print(vbCrLf + "LINT FAILED: " + LintFileShort(L));
        MsgBox(E);
        Debug.Print(E);
        Debug.Print("?LintFile(\"" + L + "\")");
        return LintFileList;

      }
    } else {
      Debug.PrintNNL(Switch(Right(L, 3) == "frm", "o", Right(L, 3) == "cls", "x", true, "."));
    }
    X = X + 1;
    if (X >= lintDotsPerRow) {
      X = 0;
      Console.WriteLn();
    }
    DoEvents();
  }
  Debug.Print(vbCrLf + "Done (" + DateDiff("s", StartTime, DateTime.Now;) + "secs).");
  LintFileList = true;
  return LintFileList;
}

public static bool LintFile(string FileName, ref string ErrStr, bool AutoFix= false) {
  bool LintFile = false;
  bool Alert = false;
  bool aOutput = false;

  Alert = ErrStr == "#";
  aOutput = ErrStr == ".";
  ErrStr = "";
  LintFile = true;

//  FileName = MakePathAbsolute(FileName, DevelopmentFolder)
  if (!FileExists(FileName())) {
    LintFile = true;
    return LintFile;

  }
  if (CheckNoLint(FileName())) {
    LintFile = true;
    return LintFile;

  }

  LintFile = LintFile && LintFileOptions(FileName(), ref ErrStr);
  LintFile = LintFile && LintFileIndent(FileName(), ref ErrStr, AutoFix);
  LintFile = LintFile && LintFileNaming(FileName(), ErrStr, AutoFix);
  LintFile = LintFile && LintFileControlNaming(FileName(), ErrStr, AutoFix);
  LintFile = LintFile && LintFileBadCode(FileName(), ref ErrStr, AutoFix);

  if (AutoFix) { // Re-run to test after Auto-fix
    ErrStr = "";
    LintFile = LintFile[FileName, ErrStr];
  }

  if (ErrStr != "") {
    if (aOutput) {
      Debug.Print(ErrStr);
    }
    if (Alert) {
      MsgBox(ErrStr);
    }
  } else {
    LintFile = true;
  }
  return LintFile;
}

private static int LintModuleFirstLine(string FileName) {
  int LintModuleFirstLine = 0;
  string S = "";
  int N = 0;
  string K = "";

  S = ReadEntireFile(FileName());
  S = Left(S, InStr(S, "Attribute VB_Name"));
  LintModuleFirstLine = CountLines(S, false, false);

  do {
    K = ReadFile(FileName(), LintModuleFirstLine, 1);
    if (!LMatch(K, "Attribute ")) {
      return LintModuleFirstLine;

    }
    if (K == "") {
      return LintModuleFirstLine;

    }
    LintModuleFirstLine = LintModuleFirstLine + 1;
  } while(!(true);
  return LintModuleFirstLine;
}

public static bool LintFileOptions(string FileName, ref string ErrStr) {
  bool LintFileOptions = false;
  int I = 0;
  string L = "";
  int A = 0;
  string F = "";

  bool oExplicit = false;


  LintFileOptions = true;

  A = LintModuleFirstLine(FileName());
  for(I=A; I<A + lintTag_ScanRange; I++) {
    L = ReadFile(FileName(), I, 1);
    if (LMatch(L, lintFile_Option)) {
      F = Mid(L, Len(lintFile_Option) + 1);
      if (F == lintFile_Option_Explicit) {
        oExplicit = true;
      } else {
        AddErrStr(ErrStr, FileName(), I - A + 1, L, "Prohibited Flag: Option " + F, lintErrorTypes.ltOptn);
        LintFileOptions = false;
      }
    }
  }

  if (!oExplicit) {
    AddErrStr(ErrStr, FileName(), 1, "", "Missing Flag: Option Explicit", lintErrorTypes.ltOptn);
    LintFileOptions = false;
  }
  return LintFileOptions;
}

private static string AutoFixInit(string FileName) {
  string AutoFixInit = "";
  int A = 0;
  string FL = "";

  A = LintModuleFirstLine(FileName());
  AutoFixInit = DevelopmentFolder() + "templint.txt";
  FL = ReadFile(FileName(), 1, A - 1);
  WriteFile(AutoFixInit, FL, true);
  return AutoFixInit;
}

private static string AutoFixLine(string FixFile, string Line, string LineFixes) {
  string AutoFixLine = "";
  dynamic FixL = null;
  dynamic KSpl = null;


  AutoFixLine = Line;
  if (LineFixes != "") {
    foreach(var FixL in Split(LineFixes, lintFixList_Sep)) {
      KSpl = Split(FixL, lintFixList_Div);
      if (KSpl(0) == "^") {
        AutoFixLine = KSpl(1) + AutoFixLine;
      }
      if (KSpl(0) == "$") {
        AutoFixLine = AutoFixLine + KSpl(1);
      }
//    If KSpl(0) = "#" Then Exit Function ' suppress output
      AutoFixLine = Replace(AutoFixLine, KSpl(0), KSpl(1));
    }
  }
  WriteFile(FixFile, AutoFixLine);
  return AutoFixLine;
}

private static string AddLineFixes(string LineFixes, string Find, string Repl) {
  string AddLineFixes = "";
  AddLineFixes = LineFixes + IIf(Len(LineFixes) == 0, "", lintFixList_Sep) + Find + lintFixList_Div + Repl;
  return AddLineFixes;
}

private static void AutoFixFinalize(string FileName, string FixFile) {
  string Contents = "";

  Contents = ReadEntireFileAndDelete(FixFile);
  while(Right(Contents, 1) == vbLf || Right(Contents, 1) == vbCr) {
    Contents = Left(Contents, Len(Contents) - 1);
  }
  Contents = Contents + vbCrLf;
  WriteFile(FileName(), Contents, true);
}

public static bool LintFileIndent(string FileName, ref string ErrStr, bool AutoFix= false) {
  bool LintFileIndent = false;
  int A = 0;

  int N = 0;
  int I = 0;

  int Continued = 0;

  int Idnt = 0;
  string Context = "";

  string OL = "";
  string L = "";
  int LNo = 0;
  string tL = "";
  string FL = "";

  int Blanks = 0;

  string FixFile = "";

  string LineFixes = "";


  if (!FileExists(FileName())) {
    LintFileIndent = true;
    return LintFileIndent;

  }
  // TODO (not supported): On Error GoTo FailedLint

  N = CountFileLines(FileName());
  A = LintModuleFirstLine(FileName());
  if (AutoFix) {
    FixFile = AutoFixInit(FileName());
  }

  for(I=A; I<N; I++) {
    L = ReadFile(FileName(), I, 1);
    OL = L;
    FL = L;
    if (Trim(L) == "") {
      Blanks = Blanks + 1;
    }
    L = DeComment(L);
    tL = LTrim(L);
    LineFixes = "";
    if (LMatch(L, "Attribute ")) {
goto NotRealLine;
    }
    LNo = I - A + 1;
//    If IsDevelopment And LNo > 275 Then Stop

    if (Trim(L) == "") {
      if (Blanks == lintLint_MaxBlankLines + 1) {
        AddErrStr(ErrStr, FileName(), LNo, OL, "Too many sequential blank lines.", lintErrorTypes.ltWhtS);
      }
goto ;
    }
    if (Continued) {
goto SkipLine;
    }

    Blanks = 0;

    if (Len(L) == Len(tL) && Right(L, 1) == ":") {
goto SkipLine; // Goto Marks
    }
    if (LMatch(tL, "On Error ")) {
goto SkipLine; // Error Handlers
    }
    if (LMatch(tL, "Debug.")) {
goto SkipLine; // Error Handlers
    }
    if (LMatch(tL, "ActiveLog")) {
goto SkipLine; // Active Logging
    }
    if (Left(tL, 1) == "#") {
goto SkipLine; // Processer Directives
    }

    if (LMatch(tL, "End Select")) {
      if (IndentContext(Context) == "Select Case Item") {
        AddIndent(Idnt, Context);
      }
      AddIndent(Idnt, Context);
    } else if (LMatch(tL, "End ") || LMatch(tL, "ElseIf ") || LMatch(tL, "Else") && !LMatch(tL, "Else ") || IsIn(IndentContext(Context), "For Loop", "For Each Loop") && tL == "Next" || LMatch(tL, "Next ") || IndentContext(Context) == "Do While Loop"& LMatch(tL, "Loop") || IndentContext(Context) == "Do Until Loop" && LMatch(tL, "Loop") || IndentContext(Context) == "Do Loop" && LMatch(tL, "Loop")) {
      if (!AddIndent(Idnt, Context, _, true)) {
        AddErrStr(ErrStr, FileName(), LNo, OL, "Cannot set negative indent.", lintErrorTypes.ltIdnt);
      }
    } else if (LMatch(tL, "Case ")) {
      if (IndentContext(Context) == "Select Case Item") {
        AddIndent(Idnt, Context);
      }
    }

//If LNo >= 383 Then Stop
//If InStr(FileName, "Functions") Then Stop
//If IsInStr(tL, "Property") Then Stop
    if (Idnt != (Len(L) - Len(tL))) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "Expected Indent " + Idnt + ", is " + (Len(L) - Len(tL)) + ": " + IndentContext(Context), lintErrorTypes.ltIdnt);
      FL = Space(Idnt) + LTrim(OL);
    }

    if (LMatch(DeString(tL), "Declare ")) {
// ignore API declarations for now
    } else if (LMatch(tL, "Function ")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "Function should be declared either Public or Private.  Neither specified.", lintErrorTypes.ltDECL);
      if (IsNotInStr(DeSpace(L), ": End ")) {
        AddIndent(Idnt, Context, "Function");
      }
      LineFixes = AddLineFixes(LineFixes, "^", "Public ");
    } else if (LMatch(tL, "Sub ")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "Sub should be declared either Public or Private.  Neither specified.", lintErrorTypes.ltDECL);
      if (IsNotInStr(DeSpace(L), ": End ")) {
        AddIndent(Idnt, Context, "Sub");
      }
      LineFixes = AddLineFixes(LineFixes, "^", "Public ");
    } else if (LMatch(tL, "Property ")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "Property should be declared either Public or Private.  Neither specified.", lintErrorTypes.ltDECL);
      if (IsNotInStr(DeSpace(L), ": End ")) {
        AddIndent(Idnt, Context, "Property");
      }
      LineFixes = AddLineFixes(LineFixes, "^", "Public ");
    } else if (LMatch(tL, "Private Function ") || LMatch(tL, "Private Sub ") || LMatch(tL, "Private Property ") || LMatch(tL, "Public Function ") || LMatch(tL, "Public Sub ") || LMatch(tL, "Public Property ") || LMatch(tL, "Friend Function ") || LMatch(tL, "Friend Sub ") || LMatch(tL, "Friend Property ")) {
      if (IsNotInStr(DeSpace(L), ": End ")) {
        AddIndent(Idnt, Context, SplitWord(tL, 2));
      }
    } else if (LMatch(tL, "For Each ")) {
      if (IsNotInStr(DeSpace(L), ": Next")) {
        AddIndent(Idnt, Context, "For Each Loop");
      }
    } else if (LMatch(tL, "For ")) {
      if (IsNotInStr(DeSpace(L), ": Next")) {
        AddIndent(Idnt, Context, "For Loop");
      }
    } else if (LMatch(tL, "While ")) {
      if (IsNotInStr(DeSpace(L), ": Loop")) {
        AddIndent(Idnt, Context, "While Loop");
      }
    } else if (tL == "Do") {
      if (IsNotInStr(DeSpace(L), ": Loop")) {
        AddIndent(Idnt, Context, "Do Loop");
      }
    } else if (LMatch(tL, "Do While ")) {
      if (IsNotInStr(DeSpace(L), ": Loop")) {
        AddIndent(Idnt, Context, "Do While Loop");
      }
    } else if (LMatch(tL, "Do Until ")) {
      if (IsNotInStr(DeSpace(L), ": Loop")) {
        AddIndent(Idnt, Context, "Do Until Loop");
      }
    } else if (LMatch(tL, "With ")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "WITH Deprecated--unsupported in all upgrade paths.", lintErrorTypes.ltWITH);
      if (IsNotInStr(L, "End With")) {
        AddIndent(Idnt, Context, "With Block");
      }
    } else if (LMatch(tL, "Select Case ")) {
      AddIndent(Idnt, Context, "Select Block");
    } else if (LMatch(tL, "Case ")) {
//      If IndentContext(Context) = "Select Case Item" Then AddIndent Idnt, Context, , True
      if (IsNotInStr(tL, ": ")) {
        AddIndent(Idnt, Context, "Select Case Item");
      }
    } else if ((LMatch(tL, "Type ") || LMatch(tL, "Private Type ") || LMatch(tL, "Public Type ")) && !LMatch(tL, "Type As ")) {
      if (IsNotInStr(L, "End Type")) {
        AddIndent(Idnt, Context, "Type Def");
      }
    } else if (LMatch(tL, "Enum ") || LMatch(tL, "Private Enum") || LMatch(tL, "Public Enum")) {
      if (IsNotInStr(L, "End Enum")) {
        AddIndent(Idnt, Context, "Enum");
      }
    } else if (LMatch(tL, "If ")) {
      if (Right(tL, 5) == " Then" || Right(tL, 2) == " _") {
        AddIndent(Idnt, Context, "If Block");
      }
    } else if (LMatch(tL, "Else") && !LMatch(tL, "Else ")) {
      AddIndent(Idnt, Context, "Else Block");
    } else if (LMatch(tL, "ElseIf ")) {
      AddIndent(Idnt, Context, "ElseIf Block");
    }

    if (IsInStr(DeString(tL), "Wend")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "WEND is deprecated.  Use Do While X ... Loop or Do ... Loop While X", lintErrorTypes.ltDEPR);
    } else if (IsInStr(" " + DeString(tL), " Next ")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "NEXT no longer needs its operand.  Remove Variable name after Next.", lintErrorTypes.ltDEPR);
    } else if (IsInStr(" " + DeString(tL), " Call ")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "CALL is no longer required.  Do not use CALL keyword in code.", lintErrorTypes.ltDEPR);
      LineFixes = AddLineFixes(LineFixes, "Call ", "");
    } else if (IsInStr(DeString(tL), "GoSub")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "GOSUB is deprecated and should not be used.", lintErrorTypes.ltDEPR);
    } else if (IsInStr(DeString(tL), "$(")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "Type-casting functions is deprecated.  Remove $ before (...).", lintErrorTypes.ltDEPR);
      LineFixes = AddLineFixes(LineFixes, "$(", "(");
    } else if (tL == "Return") {
      AddErrStr(ErrStr, FileName(), LNo, OL, "GOSUB / RETURN is deprecated and should not be used.", lintErrorTypes.ltDEPR);
    } else if (IsInStr(DeString(tL), " Stop") && Right(tL, 4) == "Stop") {
      if (!IsInStr(tL, "IsDevelopment")) {
        AddErrStr(ErrStr, FileName(), LNo, OL, "Code contains STOP statement.", lintErrorTypes.ltSTOP);
      }
    }

SkipLine:
    Continued = (Right(L, 2) == " _");

NotRealLine:
    if (AutoFix) {
      AutoFixLine(FixFile, FL, LineFixes);
    }
  }

  if (Idnt != 0) {
    AddErrStr(ErrStr, FileName(), LNo, OL, "Indent did not close. EOF.", lintErrorTypes.ltIdnt);
  }
  if (Blanks > lintLint_MaxBlankLines_AtClose) {
    AddErrStr(ErrStr, FileName(), LNo, OL, "Too many blank lines at end of file.  Max=" + lintLint_MaxBlankLines_AtClose + ".", lintErrorTypes.ltWhtS);
  }

  if (AutoFix) {
    AutoFixFinalize(FileName(), FixFile);
  }

  return LintFileIndent;


FailedLint:
  AddErrStr(ErrStr, FileName(), LNo, "", "Lint Error", lintErrorTypes.ltLErr);
  // TODO (not supported):   Resume Next
  return LintFileIndent;
}

private static bool LintFileTestName(string dName, out string ErrStr) {
  bool LintFileTestName = false;
  LintFileTestName = false;

// TODO: This check is only to avoid the problem of Dim SomVar(0, 0) for commas embedded in var names...
  if (ReduceString(dName, STR_CHR_UCASE + STR_CHR_LCASE + "_", "", 0, false) == "") {
    LintFileTestName = true;
    return LintFileTestName;

  }

  if (dName == LCase(dName)) {
    ErrStr = "Name [" + dName + "] Is All Lower Case";
  } else if (IsIn(Right(dName, 1), "%", "&", "@", "!", "#", "$")) {
    string C = "";
    string TName = "";

//% Integer Dim L%
//& Long  Dim M&
//@ Decimal Const W@ = 37.5
//! Single  Dim Q!
//# Double  Dim X#
//$ String  Dim V$ = "Secret"
    C = Right(dName, 1);
    TName = Switch(C == "%", "Long", C == "&", "Long", C == "@", "Double", C == "!", "Double", C == "#", "Double", C == "$", "String", true, "UNKNOWN-TYPE-KEY");
    ErrStr = "Type declaration by variable name not allowed. Replace " + Right(dName, 1) + " with type " + TName + ".";
  } else {
    ErrStr = "";
    LintFileTestName = true;
  }
  return LintFileTestName;
}

private static string LintStandardNaming(string vN) {
  string LintStandardNaming = "";
  switch(LCase(vN)) {
// Capitalize All
    case "nl":
      LintStandardNaming = UCase(vN);
// Capitalize Second Letter...
      break;
    case "vn":
      LintStandardNaming = LCase(Left(vN, 1)) + UCase(Mid(vN, 2, 1)) + LCase(Mid(vN, 3));
// Capitalize First Letter (default)
      break;
    default:
      LintStandardNaming = Capitalize(vN);
break;
}
  return LintStandardNaming;
}

private static bool LintFileTestArgN(string dName, ref string ErrStr) {
  bool LintFileTestArgN = false;
  LintFileTestArgN = LintFileTestName(dName, ErrStr);
  return LintFileTestArgN;
}

private static bool LintFileTestType(string DType, out string ErrStr) {
  bool LintFileTestType = false;
  LintFileTestType = true;
  switch(Trim(DType)) {
    case "Integer":
      ErrStr = "Integer should not be used here.  Use Long.";
      LintFileTestType = false;
//    Case "Single"
//      ErrStr = "Single should not be used here.  Use Double."
//      LintFileTestType = False
      break;
    case "Short":
      LintFileTestType = false;
      ErrStr = "Short should not be used here.  Use Long.";
break;
}
  return LintFileTestType;
}

private static bool LintFileIsEvent(string fName, string tL) {
  bool LintFileIsEvent = false;
  LintFileIsEvent = false;
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_") && IsInStr(tL, "Private ");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_Click");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_DblClick");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_KeyDown");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_KeyUp");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_KeyPress");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_KeyDown");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_Validate");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_GotFocus");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_LostFocus");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_QueryUnload");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_OLEDragDrop");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_OLESetData");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_RouteAfterCalculate");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_Error");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_DataArrival");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_Connect");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_Close");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_ConnectionRequest");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_SendComplete");
  LintFileIsEvent = LintFileIsEvent || IsInStr(fName, "_ZipThreadDone");
  return LintFileIsEvent;
}

private static bool LintFileNaming(string FileName, ref string ErrStr, bool AutoFix= false) {
  bool LintFileNaming = false;
  string LNo = "";

  int A = 0;
  int N = 0;
  int I = 0;
  string tE = "";

  string OL = "";
  string L = "";
  string tL = "";

  string fName = "";
  string vArgs = "";
  string AName = "";
  string vDef = "";

  bool isLet = false;
  bool isSet = false;

  string vRetType = "";

  string vName = "";
  string vType = "";

  bool Continued = false;

  dynamic Decl = null;

  string FixFile = "";
  string LineFixes = "";


  if (AutoFix) {
    FixFile = AutoFixInit(FileName());
  }

  N = CountFileLines(FileName());
  A = LintModuleFirstLine(FileName());
  for(I=A; I<N; I++) {
    OL = ReadFile(FileName(), I, 1);
    L = DeComment(OL);
    tL = LTrim(L);
    LNo = I - A + 1;
    LineFixes = "";
    if (Continued) {
goto SkipLine;
    }
//    If IsDevelopment And LNo > 1822 Then Stop
//    If LNo = 58 Then Stop
//If LNo = 147 Then Stop


    if (LMatch(tL, "Public Function ") || LMatch(tL, "Public Sub ") || LMatch(tL, "Public Property ") || LMatch(tL, "Private Function ") || LMatch(tL, "Private Sub ") || LMatch(tL, "Private Property ") || LMatch(tL, "Friend Function ") || LMatch(tL, "Friend Sub ") || LMatch(tL, "Friend Property ") || LMatch(tL, "Function ") || LMatch(tL, "Sub ") || LMatch(tL, "Property ")) {
      fName = SplitWord(tL, 1, "(");
      fName = Replace(fName, "Public ", "");
      fName = Replace(fName, "Private ", "");
      fName = Replace(fName, "Friend ", "");
      fName = Replace(fName, "Function ", "");
      fName = Replace(fName, "Sub ", "");
      fName = Replace(fName, "Property ", "");
      fName = Replace(fName, "Get ", "");
      isLet = Left(Trim(fName), 4) == "Let ";
      fName = Replace(fName, "Let ", "");
      isSet = Left(Trim(fName), 4) == "Set ";
      fName = Replace(fName, "Set ", "");

      if (!LintFileTestName(fName, tE)) {
        AddErrStr(ErrStr, FileName(), LNo, OL, tE, lintErrorTypes.ltVarN);
        LineFixes = AddLineFixes(LineFixes, " " + tE, " " + LintStandardNaming(tE));
      }

//If fName = "Form_QueryUnload" Then Stop
      if (!LintFileIsEvent(fName, tL)) {
        vRetType = SplitWord(tL, 2, ")");
        if (Left(vRetType, 3) == "As ") {
          vRetType = Mid(vRetType, 4);
          if (!LintFileTestType(vRetType, tE)) {
            AddErrStr(ErrStr, FileName(), LNo, OL, tE, lintErrorTypes.ltType);
          }
        } else {
          if (IsNotInStr(OL, "Sub ") && Right(OL, 1) != "_" && !isLet && !isSet) {
            AddErrStr(ErrStr, FileName(), LNo, OL, "No Return Type On Func/Prop", lintErrorTypes.ltNTyp);
          }
        }
        vArgs = SplitWord(DeString(tL), 1, ":");
        vArgs = SplitWord(vArgs, 2, "(", true, true);
        int MM = 0;

        if (vArgs != "") {
          MM = IIf(Right(vArgs, 2) == "()", InStrRev(vArgs, ")", Len(vArgs) - 2), InStrRev(vArgs, ")")) - 1;
        }
        if (MM >= 0) {
          vArgs = Left(vArgs, MM);
        }
        foreach(var Decl in Split(DeString(vArgs), ",")) {
          Decl = Trim(Decl);
          if (Decl == "_") {
goto IgnoreParam; // Not checking multi-line declarations for now..  Could insert in-place multi-line read..
          }

          if (LMatch(Decl, "Optional ")) {
            vDef = SplitWord(Decl, 2, " = ");
            if (vDef == "") {
              AddErrStr(ErrStr, FileName(), LNo, OL, "Parameter declared OPTIONAL but no default specified. Must specify default.", lintErrorTypes.ltNOpD);
            }
            Decl = Trim(Replace(Decl, "Optional ", ""));
          }

          if (!LMatch(Decl, "ByVal ") && !LMatch(Decl, "ByRef ") && !LMatch(Decl, "ParamArray ")) {
            AddErrStr(ErrStr, FileName(), LNo, OL, "Neither ByVal nor ByRef are specified. Must Specify one or other.", lintErrorTypes.ltDECL);
            LineFixes = AddLineFixes(LineFixes, Replace(Decl, "_", ""), "ByRef " + Replace(Decl, "_", ""));
          } else {
            Decl = Replace(Decl, "ByRef ", "");
            Decl = Replace(Decl, "ByVal ", "");
            Decl = Replace(Decl, "ParamArray ", "");
            Decl = Trim(Decl);
          }

          vName = SplitWord(Decl, 1, " As ");
          if (!LintFileTestArgN(vName, tE)) {
            AddErrStr(ErrStr, FileName(), LNo, OL, tE, lintErrorTypes.ltArgN);
          }

          vType = SplitWord(Decl, 2, " As ");
          if (vType == "") {
            AddErrStr(ErrStr, FileName(), LNo, OL, "No Param Type on Func/Sub/Prop", lintErrorTypes.ltNTyp);
          }
          if (!LintFileTestType(vType, tE)) {
            AddErrStr(ErrStr, FileName(), LNo, OL, tE, lintErrorTypes.ltType);
          }

IgnoreParam:
        }
      }
    } else if (LMatch(tL, "Private Declare ") || LMatch(tL, "Public Declare ") || LMatch(tL, "Declare ")) {
    } else if (LMatch(tL, "Dim ") || LMatch(tL, "Private ") || LMatch(tL, "Public ")) {
      vArgs = tL;
      vArgs = Replace(vArgs, "Dim ", "");
      vArgs = Replace(vArgs, "Private ", "");
      vArgs = Replace(vArgs, "Public ", "");
      vArgs = Replace(vArgs, "Const ", "");

      foreach(var Decl in Split(DeString(vArgs), ",")) {
        vName = Trim(SplitWord(Decl, 1, " As "));
        vName = Trim(SplitWord(vName, 1, " = "));
        if (!LintFileTestName(vName, tE)) {
          AddErrStr(ErrStr, FileName(), LNo, OL, tE, lintErrorTypes.ltArgN);
          LineFixes = AddLineFixes(LineFixes, vName, LintStandardNaming(vName));
        }
        if (IsNotInStr(OL, "Enum ") && IsNotInStr(OL, "Type ")) {
          vType = Trim(SplitWord(Decl, 2, " As "));
          if (!LMatch(vName, "Event ")) {
            if (vType == "") {
              AddErrStr(ErrStr, FileName(), LNo, OL, "No Type on Decl", lintErrorTypes.ltNTyp);
            }
          }
          if (!LintFileTestType(vType, tE)) {
            AddErrStr(ErrStr, FileName(), LNo, OL, tE, lintErrorTypes.ltType);
          }
        }
      }
    }
SkipLine:
    Continued = (Right(L, 2) == " _");

    if (AutoFix) {
      AutoFixLine(FixFile, OL, LineFixes);
    }
  }

  if (AutoFix) {
    AutoFixFinalize(FileName(), FixFile);
  }

  LintFileNaming = ErrStr == "";
  return LintFileNaming;
}

private static bool LintFileControlNaming(string FileName, ref string ErrStr, bool AutoFix_UNUSED= false) {
  bool LintFileControlNaming = false;
  const int MaxCtrl = 128;
  int LNo = 0;

  string Contents = "";
  int I = 0;

  string Match = "";

  int N = 0;
  int K = 0;

  string CtlName = "";
  string ErrMsg = "";

  Collection cUnique = null;
  dynamic Reported = null;


  Contents = ReadEntireFile(FileName());
  cUnique = new Collection();;

  List<dynamic> vTypes = new List<dynamic> {}; // TODO - Specified Minimum Array Boundary Not Supported:   Dim vTypes() As Variant

  vTypes = Array("CheckBox", "Command", "Option", "Frame", "Label", "TextBox", "RichTextBox", "RichTextBoxNew", "ComboBox", "ListBox", "Timer", "UpDown", "HScrollBar", "Image", "Picture", "MSFlexGrid", "DBGrid", "Line", "Shape", "DTPicker");
  for(I=LBound(vTypes); I<UBound(vTypes); I++) {
// "[^a-zA-Z]" & vTypes(I) & "[0-9*]\."
    Match = "Begin [a-zA-Z0-9]*.[a-zA-Z0-9]* " + vTypes[I] + "[0-9]*";
    if (RegExTest(Contents, Match)) {
      N = RegExCount(Contents, Match);
      for(K=0; K<N - 1; K++) {
        CtlName = RegExNMatch(Contents, Match, K);
        CtlName = Split(CtlName, " ")(2);
        CtlName = Trim(CtlName);

        // TODO (not supported):         On Error Resume Next
        Reported = "";
        Reported = cUnique.Item(CtlName);
        cUnique.Add("1", CtlName);
        // TODO (not supported):         On Error GoTo 0

        if (CtlName != "" && Reported == "") {
          ErrMsg = "Default Control Name in use: " + CtlName + ".  Rename Control.";
          AddErrStr(ErrStr, FileName(), LNo, "", ErrMsg, lintErrorTypes.ltCtlN);
        }
      }
    }
  }

  LintFileControlNaming = ErrStr == "";
  return LintFileControlNaming;
}

public static bool LintFileBadCode(string FileName, ref string ErrStr, bool AutoFix= false) {
  bool LintFileBadCode = false;
  string LNo = "";

  int A = 0;
  int N = 0;
  int I = 0;
  string tE = "";

  string OL = "";
  string L = "";
  string tL = "";

  string fName = "";
  string vArgs = "";
  string AName = "";
  string vDef = "";

  string vRetType = "";

  string vName = "";
  string vType = "";

  bool Continued = false;

  dynamic Decl = null;

  string FixFile = "";
  string LineFixes = "";


  if (AutoFix) {
    FixFile = AutoFixInit(FileName());
  }

  N = CountFileLines(FileName());
  A = LintModuleFirstLine(FileName());
  for(I=A; I<N; I++) {
    OL = ReadFile(FileName(), I, 1);
    L = DeComment(OL);
    tL = LTrim(L);
    LNo = I - A + 1;
    LineFixes = "";
    if (Continued) {
goto SkipLine;
    }

    if (RegExTest(tL, "\\.Enabled = [-0-9]")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "Property [Enabled] Should Be Boolean.  Numeric found.", lintErrorTypes.ltType);
    }
    if (RegExTest(tL, "\\.Visible = [-0-9]")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "Property [Visible] Should Be Boolean.  Numeric found.", lintErrorTypes.ltType);
    }
    if (RegExTest(" " + tL, "[^a-zA-Z0-0]Me[.][^ ]")) {
      AddErrStr(ErrStr, FileName(), LNo, OL, "Self Reference [Me.*] is unnecessary.", lintErrorTypes.ltSelf); //@NO-LINT
    }

SkipLine:
    Continued = (Right(L, 2) == " _");
  }

  LintFileBadCode = ErrStr == "";
  return LintFileBadCode;
}
}
