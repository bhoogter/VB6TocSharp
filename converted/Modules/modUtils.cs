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


static class modUtils {
// Option Explicit
public const string patToken = "([a-zA-Z_][a-zA-Z_0-9]*)";
public const string patNotToken = "([^a-zA-Z_0-9])";
public const string patTokenDot = "([a-zA-Z_.][a-zA-Z_0-9.]*)";
public const string vbCrLf2 = vbCrLf + vbCrLf;
public const string vbCrLf3 = vbCrLf + vbCrLf + vbCrLf;
public const string vbCrLf4 = vbCrLf + vbCrLf + vbCrLf + vbCrLf;
public const string STR_CHR_UCASE = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
public const string STR_CHR_LCASE = "abcdefghijklmnopqrstuvwxyz";
public const string STR_CHR_DIGIT = "1234567890";


public static bool IsInStr(string Src, string Find) {
  bool IsInStr = false;
  IsInStr = InStr(Src, Find) > 0;
  return IsInStr;
}

public static bool IsNotInStr(string S, string Fnd) {
  bool IsNotInStr = false;
  IsNotInStr = !IsInStr(S, Fnd);
  return IsNotInStr;
}

public static bool FileExists(string FN) {
  bool FileExists = false;
  FileExists = FN != "" && Dir(FN) != "";
  return FileExists;
}

public static bool DirExists(string FN) {
  bool DirExists = false;
  DirExists = FN != "" && Dir(FN, vbDirectory) != "";
  return DirExists;
}

public static string FileName(string FN) {
  string FileName = "";
  FileName = Mid(FN, InStrRev(FN, "\\") + 1);
  return FileName;
}

public static string FileBaseName(string FN) {
  string FileBaseName = "";
  FileBaseName = Left(FileName(FN), InStrRev(FileName(FN), ".") - 1);
  return FileBaseName;
}

public static string FilePath(string FN) {
  string FilePath = "";
  FilePath = Left(FN, InStrRev(FN, "\\"));
  return FilePath;
}

public static string ChgExt(string FN, string NewExt) {
  string ChgExt = "";
  ChgExt = Left(FN, InStrRev(FN, ".") - 1) + NewExt;
  return ChgExt;
}

public static string tLeft(string Str, int N) {
  string tLeft = "";
  tLeft = Left(Trim(Str), N);
  return tLeft;
}

public static string tMid(string Str, int N, int M= 0) {
  string tMid = "";
  tMid = IIf(M == 0, Mid(Trim(Str), N), Mid(Trim(Str), N, M));
  return tMid;
}

public static int StrCnt(string Src, string Str) {
  int StrCnt = 0;
  StrCnt = (Len(Src) - Len(Replace(Src, Str, ""))) / Len(Str);
  return StrCnt;
}

public static bool LMatch(string Src, string tMatch) {
  bool LMatch = false;
  LMatch = Left(Src, Len(tMatch)) == tMatch;
  return LMatch;
}

public static bool tLMatch(string Src, string tMatch) {
  bool tLMatch = false;
  tLMatch = Left(LTrim(Src), Len(tMatch)) == tMatch;
  return tLMatch;
}

public static int Px(int Twips) {
  int Px = 0;
  Px = Twips / 14;
  return Px;
}

public static string Quote(string S) {
  string Quote = "";
  Quote = "\"" + S + "\"";
  return Quote;
}

public static string AlignString(string S, int N) {
  string AlignString = "";
  AlignString = Left(S + Space(N), N);
  return AlignString;
}

public static string Capitalize(string S) {
  string Capitalize = "";
  Capitalize = UCase(Left(S, 1)) + Mid(S, 2);
  return Capitalize;
}

public static string DevelopmentFolder() {
  string DevelopmentFolder = "";
  DevelopmentFolder = App.Path + "\\";
  return DevelopmentFolder;
}

public static bool IsIDE() {
  bool IsIDE = false;
//IsIDE = False
//Exit Function

// works on a very simple princicple... debug statements don't get compiled...
  // TODO (not supported):   On Error GoTo IDEInUse
  Console.WriteLine(1/0); //division by zero error
  IsIDE = false;
  return IsIDE;

IDEInUse:;
  IsIDE = true;
  return IsIDE;
}

public static Variant) As Boolean IsIn(string S, dynamic ParamArray_UNUSED) {
  Variant) As Boolean IsIn = null;
  dynamic L = null;

  foreach(var L in K) {
    if (S == L) {
      IsIn = true;
      return IsIn;

    }
  }
  return IsIn;
}

public static bool WriteOut(string F, string S, string O= "") {
  bool WriteOut = false;
  if (!IsConverted(F, O)) {
    WriteOut = WriteFile(OutputFolder(O) + F, S, true);
  } else {
    Console.WriteLine("Already converted: " + F);
  }
  return WriteOut;
}

public static bool IsConverted(string F, string O= "") {
  bool IsConverted = false;
  IsConverted = IsInStr(Left(ReadEntireFile(OutputFolder(O) + F), 100), "### CONVERTED");
  return IsConverted;
}

public static string FileExt(string FN, bool vLCase= true) {
  string FileExt = "";
  if (FN == "") {
    return FileExt;

  }
  if (InStr(FN, ".") == 0) {
    return FileExt;

  }
  FileExt = Mid(FN, InStrRev(FN, "."));
  FileExt = IIf(vLCase, LCase(FileExt), FileExt);
  return FileExt;
}

public static string deQuote(string Src) {
  string deQuote = "";
  if (Left(Src, 1) == "\"") {
    Src = Mid(Src, 2);
  }
  if (Right(Src, 1) == "\"") {
    Src = Left(Src, Len(Src) - 1);
  }
  deQuote = Src;
  return deQuote;
}

public static string deWS(string S) {
  string deWS = "";
  while(IsInStr(S, " " + vbCrLf)) {
    S = Replace(S, " " + vbCrLf, vbCrLf);
  }
  while(IsInStr(S, vbCrLf4)) {
    S = Replace(S, vbCrLf4, vbCrLf3);
  }

  S = Replace(S, "{" + vbCrLf2, "{" + vbCrLf);
  S = RegExReplace(S, "(" + vbCrLf2 + ")([ ]*{)", vbCrLf + "$2");
  S = RegExReplace(S, "([ ]*case .*:)" + vbCrLf2, "$1" + vbCrLf);
  deWS = S;
  return deWS;
}

public static string nlTrim(string Str) {
  string nlTrim = "";
  while(InStr(" " + vbTab + vbCr + vbLf, Left(Str, 1)) != 0& Str != "") {
    Str = Mid(Str, 2);
  }
  while(InStr(" " + vbTab + vbCr + vbLf, Right(Str, 1)) != 0& Str != "") {
    Str = Mid(Str, 1, Len(Str) - 1);
  }
  nlTrim = Str;
  return nlTrim;
}

public static string sSpace(int N) {
  string sSpace = "";
  // TODO (not supported): On Error Resume Next
  sSpace = Space(N);
  return sSpace;
}

public static string nextBy(string Src, string Del= "\"", int Ind= 1, bool ProcessVBComments_UNUSED= false) {
  string nextBy = "";
  int L = 0;

  DoEvents();
  L = InStr(Src, Del);
  if (L == 0) {
    nextBy = IIf(Ind <= 1, Src, "");
    return nextBy;

  }
  if (Ind <= 1) {
    nextBy = Left(Src, L - 1);
  } else {
    nextBy = nextBy[Mid(Src, L + Len(Del)), Del, Ind - 1];
  }
  return nextBy;
}

public static int StrQCnt(string Src, string Str) {
  int StrQCnt = 0;
  int N = 0;
  int I = 0;
  string C = "";

  bool Q = false;


  StrQCnt = 0;
  N = Len(Src);
  for(I=1; I<N; I++) {
    C = Mid(Src, I, 1);
    if (C == "\"") {
      Q = !Q;
    } else {
      if (!Q) {
        if (LMatch(Mid(Src, I), Str)) {
          StrQCnt = StrQCnt + 1;
        }
      }
    }
  }
  return StrQCnt;
}

public static int nextByPCt(string Src, string Del= "\"", int Ind_UNUSED= 1) {
  int nextByPCt = 0;
  int M = 0;
  int N = 0;
  string F = "";

  N = 0;
  do {
    N = N + 1;
    if (N > 1000) {
      break;
    }
    F = nextByP(Src, Del, N);
    if (F == "") {
      M = M + 1;
      if (M >= 10) {
        break;
      }
    } else {
      M = 0;
    }
  } while(!(true));
  nextByPCt = N - M;
  return nextByPCt;
}

public static string nextByP(string Src, string Del= "\"", int Ind= 1) {
  string nextByP = "";
  string F = "";
  int N = 0;
  int M = 0;

  string R = "";
  string T = "";

  N = 0;
  F = "";
  do {
    M = M + 1;
    if (M > 100) {
      break;
    }
    N = N + 1;
    T = nextBy(Src, Del, N);
    R = R + IIf(Len(R) == 0, "", Del) + T;
  } while(!(StrQCnt(R, "(") == StrQCnt(R, ")")));
  if (Ind <= 1) {
    nextByP = R;
  } else {
    nextByP = nextByP[Mid(Src, Len(R) + Len(Del) + 1), Del, Ind - 1];
  }
  return nextByP;
}

public static string NextByOp(string Src, int Ind= 1, out string Op) {
  string NextByOp = "";
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

  P() = A;
  K = 3;
  if (Len(P()) > Len(S)) {
    P() = S;
    K = 3;
  }
  if (Len(P()) > Len(M)) {
    P() = M;
    K = 3;
  }
  if (Len(P()) > Len(D)) {
    P() = D;
    K = 3;
  }
  if (Len(P()) > Len(I)) {
    P() = I;
    K = 3;
  }
  if (Len(P()) > Len(C)) {
    P() = C;
    K = 3;
  }
  if (Len(P()) > Len(E)) {
    P() = E;
    K = 3;
  }

  if (Len(P()) > Len(cNE)) {
    P() = cNE;
    K = 4;
  }
  if (Len(P()) > Len(cLT)) {
    P() = cLT;
    K = 3;
  }
  if (Len(P()) > Len(cGT)) {
    P() = cGT;
    K = 3;
  }
  if (Len(P()) > Len(cLE)) {
    P() = cLE;
    K = 4;
  }
  if (Len(P()) > Len(cGE)) {
    P() = cGE;
    K = 4;
  }
  if (Len(P()) > Len(cEQ)) {
    P() = cEQ;
    K = 3;
  }

  if (Len(P()) > Len(lA)) {
    P() = lA;
    K = 5;
  }
  if (Len(P()) > Len(lO)) {
    P() = lO;
    K = 4;
  }
  if (Len(P()) > Len(lM)) {
    P() = lM;
    K = 5;
  }
  if (Len(P()) > Len(LL)) {
    P() = LL;
    K = 6;
  }

  if (Len(P()) > Len(xLk)) {
    P() = xLk;
    K = 6;
  }
  if (Len(P()) > Len(xIs)) {
    P() = xIs;
    K = 4;
  }

  NextByOp = P();
  if (Ind <= 1) {
    Op = Mid(Src, Len(P()) + 1, K);
    NextByOp = P();
  } else {
    NextByOp = NextByOp[Trim(Mid(Src, Len(P) + 3)), Ind - 1, Op];
  }
  return NextByOp;
}

public static string ReplaceToken(string Src, string OrigToken, string NewToken) {
  string ReplaceToken = "";
  ReplaceToken = RegExReplace(Src, "([^a-zA-Z_0-9])(" + OrigToken + ")([^a-zA-Z_0-9])", "$1" + NewToken + "$3");
  return ReplaceToken;
}

public static string SplitWord(string Source, int N= 1, string Space= " ", bool TrimResult= true, bool IncludeRest= false) {
  string SplitWord = "";

//::::SplitWord
//:::SUMMARY
//: Return an indexed word from a string
//:::DESCRIPTION
//: Split()s a string based on a space (or other character) and return the word specified by the index.
//: - Returns "" for 1 > N > Count
//:::PARAMETERS
//: - Source - The original source string to analyze
//: - [N] = 1 - The index of the word to return (Default = 1)
//: - [Space] = " " - The character to use as the "space" (defaults to %20).
//: - [TrimResult] - Apply Trim() to the result (Default = True)
//: - [IncludeRest] - Return the rest of the string starting at the indexed word (Default = False).
//:::EXAMPLE
//: - SplitWord("The Rain In Spain Falls Mostly", 4) == "Spain"
//: - SplitWord("The Rain In Spain Falls Mostly", 4, , , True) == "Spain Falls Mostly"
//: - SplitWord("a:b:c:d", -1, ":") === "d"
//:::RETURN
//:  String
//:::SEE ALSO
//: Split, CountWords
  dynamic S = null;
  int I = 0;

  N = N - 1;
  if (Source == "") {
    return SplitWord;

  }
  S = Split(Source, Space);
  if (N < 0) {
    N = UBound(S) + N + 2;
  }
  if (N < LBound(S) || N > UBound(S)) {
    return SplitWord;

  }
  if (!IncludeRest) {
    SplitWord = S(N);
  } else {
    for(I=N; I<UBound(S); I++) {
      SplitWord = SplitWord + IIf(Len(SplitWord) > 0, Space, "") + S(I);
    }
  }
  if (TrimResult) {
    SplitWord = Trim(SplitWord);
  }
  return SplitWord;
}

public static int CountWords(string Source_UNUSED, string Space_UNUSED= " ") {
  int CountWords = 0;
//::::CountWords
//:::SUMMARY
//: Returns the number of words in a string (determined by <Space> parameter)
//:::DESCRIPTION
//: Returns the count of words.
//:::PARAMETERS
//: - Source - The original source string to analyze
//: - [Space] = " " - The character to use as the "space" (defaults to %20).
//:::EXAMPLE
//: - CountWords("The Rain In Spain Falls Mostly") == 6
//: - CountWords("The Rain In Spain Falls Mostly", "n") == 4
//:::RETURN
//:  String
//:::SEE ALSO
//: SplitWord
  dynamic L = null;

// Count actual words.  Blank spaces don't count, before, after, or in the middle.
// Only a simple split and loop--there may be faster ways...
  foreach(var L in Split(Source, Space)) {
    if (L != "") {
      CountWords = CountWords + 1;
    }
  }
  return CountWords;
}

public static dynamic ArrSlice(ref dynamic sourceArray, int fromIndex, int toIndex) {
  dynamic ArrSlice = null;
  int Idx = 0;

  List<dynamic> tempList = new List<dynamic> {}; // TODO - Specified Minimum Array Boundary Not Supported:   Dim tempList()


  if (!IsArray(sourceArray)) {
    return ArrSlice;

  }

  fromIndex = FitRange(LBound(sourceArray), fromIndex, UBound(sourceArray));
  toIndex = FitRange(fromIndex, toIndex, UBound(sourceArray));

  for(Idx=fromIndex; Idx<toIndex; Idx++) {
    ArrAdd(tempList, sourceArray[Idx]);
  }

  ArrSlice = tempList;
  return ArrSlice;
}

public static void ArrAdd(ref dynamic Arr(_UNUSED) {
  int X = 0;

  Err().Clear();
  // TODO (not supported): On Error Resume Next
  X = UBound(Arr);
  if (Err().Number != 0) {
    Arr = Array(Item);
return;

  }
  List<> Arr_8186_tmp = new List<>();
for (int redim_iter_2138=0;i<0;redim_iter_2138++) {Arr.Add(redim_iter_2138<Arr.Count ? Arr(redim_iter_2138) : null);}
  Arr(UBound(Arr)) = Item;
}

public static dynamic SubArr(dynamic sourceArray, int fromIndex, int copyLength) {
  dynamic SubArr = null;
  SubArr = ArrSlice(ref sourceArray, fromIndex, fromIndex + copyLength - 1);
  return SubArr;
}

public static bool InRange(dynamic LBnd, dynamic CHK, dynamic UBnd, bool IncludeBounds= true) {
  bool InRange = false;
// TODO (not supported): On Error Resume Next // because we're doing this as variants..
  if (IncludeBounds) {
    InRange = (CHK >= LBnd) && (CHK <= UBnd);
  } else {
    InRange = (CHK > LBnd) && (CHK < UBnd);
  }
  return InRange;
}

public static dynamic FitRange(dynamic LBnd, dynamic CHK, dynamic UBnd) {
  dynamic FitRange = null;
  // TODO (not supported): On Error Resume Next
  if (CHK < LBnd) {
    FitRange = LBnd;
  } else if (CHK > UBnd) {
    FitRange = UBnd;
  } else {
    FitRange = CHK;
  }
  return FitRange;
}

public static int CodeSectionLoc(string S) {
  int CodeSectionLoc = 0;
  const string Token = "Attribute VB_Name";
  int N = 0;
  int K = 0;


  N = InStr(S, Token);
  if (N == 0) {
    return CodeSectionLoc;

  }
  do {
    N = InStr(N, S, vbLf) + 1;
    if (N <= 1) {
      return CodeSectionLoc;

    }
  } while(!(Mid(S, N, 10) == "Attribute "));

  CodeSectionLoc = N;
  return CodeSectionLoc;
}

public static dynamic CodeSectionGlobalEndLoc(string S) {
  dynamic CodeSectionGlobalEndLoc = null;
  do {
    CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc + RegExNPos(Mid(S, CodeSectionGlobalEndLoc + 1), "([^a-zA-Z0-9_]Function |[^a-zA-Z0-9_]Sub |[^a-zA-Z0-9_]Property )") + 1;
    if (CodeSectionGlobalEndLoc == 1) {
      CodeSectionGlobalEndLoc = Len(S);
      return CodeSectionGlobalEndLoc;

    }
  } while(!(Mid(S, CodeSectionGlobalEndLoc - 8, 8) == "Declare "));
  if (CodeSectionGlobalEndLoc >= 8) {
    if (Mid(S, CodeSectionGlobalEndLoc - 7, 7) == "Friend ") {
      CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc - 7;
    }
    if (Mid(S, CodeSectionGlobalEndLoc - 7, 7) == "Public ") {
      CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc - 7;
    }
    if (Mid(S, CodeSectionGlobalEndLoc - 8, 8) == "Private ") {
      CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc - 8;
    }
  }
  CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc - 1;
  return CodeSectionGlobalEndLoc;
}

public static bool isOperator(string S) {
  bool isOperator = false;
  switch(Trim(S)) {
    case "+":
      isOperator = true;
      break;
    default:
      isOperator = false;
break;
}
  return isOperator;
}

public static void Prg(int Val= -1, int Max= -1, dynamic Cap= "#") {
  frm.Prg(Val, Max, Cap);
}

public static string cVal(ref Collection Coll, ref string Key, string Def= "") {
  string cVal = "";
  // TODO (not supported):   On Error Resume Next
  cVal = Def;
  cVal = Coll.Item(LCase(Key));
  return cVal;
}

public static string cValP(ref Collection Coll, ref string Key, string Def= "") {
  string cValP = "";
  cValP = P(deQuote(cVal(ref Coll, ref Key, Def)));
  return cValP;
}

public static string P(string Str) {
  string P = "";
  Str = Replace(Str, "&", "&amp;");
  Str = Replace(Str, "<", "&lt;");
  Str = Replace(Str, ">", "&gt;");
  P = Str;
  return P;
}

public static string ModuleName(string S) {
  string ModuleName = "";
  int J = 0;
  int K = 0;

  const string NameTag = "Attribute VB_Name = \"";
  J = InStr(S, NameTag) + Len(NameTag);
  K = InStr(J, S, "\"") - J;
  ModuleName = Mid(S, J, K);
  return ModuleName;
}

public static dynamic IsInCode(string Src, int N_UNUSED) {
  dynamic IsInCode = null;
  int I = 0;
  string C = "";

  bool Qu = false;

  IsInCode = false;
  for(I=N; I<1 Step -1; I++) {
    C = Mid(Src, I, 1);
    if (C == vbCr || C == vbLf) {
      IsInCode = true;
      return IsInCode;

    } else if (C == "\"") {
      Qu = !Qu;
    } else if (C == "'") {
      if (!Qu) {
        return IsInCode;

      }
    }
  }
  IsInCode = true;
  return IsInCode;
}

public static string TokenList(string S) {
  string TokenList = "";
  int I = 0;
  int N = 0;
  string T = "";

  N = RegExCount(S, patToken);
  for(I=0; I<N - 1; I++) {
    T = RegExNMatch(S, patToken, I);
    TokenList = TokenList + "," + T;
  }
  return TokenList;
}

public static int Random(int Max= 10000) {
  int Random = 0;
  Randomize();
  Random = ((Rnd * Max) + 1);
  return Random;
}

public static string Stack(ref string Src, string Val= "##REM##", bool Peek= false) {
  string Stack = "";
  if (Val == "##REM##") {
    Stack = nextBy(Src, ",");
    if (!Peek) {
      Src = Mid(Src, Len(Stack) + 2);
    }
    Stack = Replace(Stack, "\"\"", "\"");
    if (Left(Stack, 1) == "\"") {
      Stack = Mid(Stack, 2);
      Stack = Left(Stack, Len(Stack) - 1);
    }
  } else {
    Src = "\"" + Replace(Val, "\"", "\"\"") + "\"," + Src;
    Stack = Val;
  }
  return Stack;
}

public static string QuoteXML(string S) {
  string QuoteXML = "";
  QuoteXML = S;
  QuoteXML = Replace(S, "\"", "&quot;");
  QuoteXML = Quote(QuoteXML);
  return QuoteXML;
}

public static string ReduceString(string Src, string Allowed= "", string Subst= "-", int MaxLen= 0, bool bLCase= true) {
  string ReduceString = "";
//::::ReduceString
//:::SUMMARY
//: Reduces a string by removing non-allowed characters, optionally replacing them with a substitute.
//:::DESCRIPTION
//: Non-allowed characters are removed, and, if supplied, replaced with a substitute.
//: Substitutes are trimmed from either end, and all duplicated substitutes are remvoed.
//:
//: After this process, the string can be given LCase (default) or truncated (not default), if desired.
//:
//: This is effectively a slug maker, although it is somewhat adaptable to any cleaning routine.
//:::PARAMETERS
//: - Src - Source string to be reduced
//: - [Allowed] - The list of allowable characters.  Defaults to [A-Za-z0-9]*
//: - [Subst] - If specified, the character to replace non-allowed characters with (default == "-")
//: - [MaxLen] - If passed, truncates longer strings to this length.  Default = 0
//: - [bLCase] - Convert string to lower case after operation.  Default = True
//:::EXAMPLE
//: - ReduceString("   Something To be 'slugified'!!!****") == "something-to-be-slugified"
//:::RETURN
//:  String - The slug generated from the source.
//:::AUTHOR
//: Benjamin - 2018.04.28
//:::SEE ALSO
//:  ArrangeString, StringNumerals, slug, CleanANI
  int I = 0;
  int N = 0;
  string C = "";

  if (Allowed == "") {
    Allowed = STR_CHR_UCASE + STR_CHR_LCASE + STR_CHR_DIGIT;
  }
  ReduceString = "";
  N = Len(Src);
  for(I=1; I<N; I++) {
    C = Mid(Src, I, 1);
    ReduceString = ReduceString + IIf(IsInStr(Allowed, C), C, Subst);
  }

  if (Subst != "") {
    while(IsInStr(ReduceString, Subst + Subst)) {
      ReduceString = Replace(ReduceString, Subst + Subst, Subst);
    }
    while(Left(ReduceString, Len(Subst)) == Subst) {
      ReduceString = Mid(ReduceString, Len(Subst) + 1);
    }
    while(Right(ReduceString, Len(Subst)) == Subst) {
      ReduceString = Left(ReduceString, Len(ReduceString) - Len(Subst));
    }
  }

  if (MaxLen > 0) {
    ReduceString = Left(ReduceString, MaxLen);
  }
  if (bLCase) {
    ReduceString = LCase(ReduceString);
  }
  return ReduceString;
}
}
