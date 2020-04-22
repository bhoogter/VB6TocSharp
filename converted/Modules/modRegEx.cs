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


static class modRegEx {
// Option Explicit
private static dynamic mRegEx = null;


static dynamic RegEx {
  get {
    dynamic RegEx;
    if (mRegEx == null) {
  mRegEx = CreateObject("vbscript.regexp");
  mRegEx.Global = true;
}
RegEx = mRegEx;

  return RegEx;
  }
}


public static bool RegExTest(string Src, string Find) {
  bool RegExTest = false;
  // TODO (not supported): On Error Resume Next
  RegEx.Pattern = Find;
  RegExTest = RegEx.Test(Src);
  return RegExTest;
}

public static int RegExCount(string Src, string Find) {
  int RegExCount = 0;
  // TODO (not supported): On Error Resume Next
  RegEx.Pattern = Find;
  RegEx.Global = true;
  RegExCount = RegEx.Execute(Src).Count;
  return RegExCount;
}

public static int RegExNPos(string Src, string Find, int N= 0) {
  int RegExNPos = 0;
  // TODO (not supported): On Error Resume Next
  dynamic RegM = null;
  string tempStr = "";
  string tempStr2 = "";

  RegEx.Pattern = Find;
  RegEx.Global = true;
  RegExNPos = RegEx.Execute(Src).Item(N).FirstIndex + 1;
  return RegExNPos;
}

public static string RegExNMatch(string Src, string Find, int N= 0) {
  string RegExNMatch = "";
  // TODO (not supported): On Error Resume Next
  dynamic RegM = null;
  string tempStr = "";
  string tempStr2 = "";

  RegEx.Pattern = Find;
  RegEx.Global = true;
  RegExNMatch = RegEx.Execute(Src).Item(N).Value;
  return RegExNMatch;
}

public static string RegExReplace(string Src, string Find, string Repl) {
  string RegExReplace = "";
  // TODO (not supported): On Error Resume Next
  dynamic RegM = null;
  string tempStr = "";
  string tempStr2 = "";

  RegEx.Pattern = Find;
  RegEx.Global = true;
  RegExReplace = RegEx.Replace(Src, Repl);
  return RegExReplace;
}

public static dynamic RegExSplit(string szStr, string szPattern) {
  dynamic RegExSplit = null;
  // TODO (not supported): On Error Resume Next
  dynamic oAl = null;
  dynamic oRe = null;
  dynamic oMatches = null;

  oRe = RegEx;
  oRe.Pattern = "^(.*)(" + szPattern + ")(.*)$";
  oRe.IgnoreCase = true;
  oRe.Global = true;
  oAl = CreateObject("System.Collections.ArrayList");

  do {
    oMatches = oRe.Execute(szStr);
    if (oMatches.Count > 0) {
      oAl.Add(oMatches(0).SubMatches(2));
      szStr = oMatches(0).SubMatches(0);
    } else {
      oAl.Add(szStr);
      break;
    }
  }
  oAl.Reverse();
  RegExSplit = oAl.ToArray;
  return RegExSplit;
}

public static int RegExSplitCount(string szStr, string szPattern) {
  int RegExSplitCount = 0;
  // TODO (not supported): On Error Resume Next
  List<dynamic> T = new List<dynamic> {}; // TODO - Specified Minimum Array Boundary Not Supported:   Dim T()

  T = RegExSplit(szStr, szPattern);
  RegExSplitCount = UBound(T) - LBound(T) + 1;
  return RegExSplitCount;
}
}
