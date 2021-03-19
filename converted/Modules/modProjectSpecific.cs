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


static class modProjectSpecific {
// Option Explicit


public static string ProjectSpecificPostCodeLineConvert(string Str) {
  string ProjectSpecificPostCodeLineConvert = "";
  string S = "";

  S = Str;

//  If IsInStr(S, "!C == null") Then Stop

// Some patterns we dont use or didn't catch in lint...
  if (IsInStr(S, "DisposeDA")) {
    S = Replace(S, "DisposeDA", "// DisposeDA");
  }
  if (IsInStr(S, "MousePointer = vbNormal")) {
    S = Replace(S, "MousePointer = vbNormal", "MousePointer = vbDefault");
  }

// We use decimal, not double
  if (IsInStr(S, "Val(")) {
    S = Replace(S, "Val( ", "ValD(");
  }

// Bad pattern combination
  if (RegExTest(S, "\\(!" + patToken + " == null\\)")) {
    S = Replace(S, "!", "", 1);
    S = Replace(S, "==", "!=", 1);
  }

// False ref entries...
  if (IsInStr(S, "IsIn(")) {
    S = Replace(S, "ref ", "");
  }
  if (IsInStr(S, "POMode(")) {
    S = Replace(S, "ref ", "");
  }
  if (IsInStr(S, "OrderMode(")) {
    S = Replace(S, "ref ", "");
  }
  if (IsInStr(S, "InvenMode(")) {
    S = Replace(S, "ref ", "");
  }
  if (IsInStr(S, "ReportsMode(")) {
    S = Replace(S, "ref ", "");
  }
  if (IsInStr(S, "SetButtonImage(")) {
    S = Replace(S, "ref ", "");
    S = Replace(S, ".DefaultProperty", "");
  }
  if (IsInStr(S, "EnableFrame")) {
    S = Replace(S, "ref ", "");
  }
  S = Replace(S, " && BackupType.", " & BackupType.");

// Common Mistake Functions...
  if (IsInStr(S, "StoreSettings.")) {
    S = Replace(S, "StoreSettings.", "StoreSettings().");
  }

// etc
  if (IsInStr(S, ".hwnd")) {
    S = Replace(S, ".hwnd", ".hWnd()");
  }
  if (IsInStr(S, "SetCustomFrame")) {
    S = "";
  }
  if (IsInStr(S, "RemoveCustomFrame")) {
    S = "";
  }
  S = Replace(S, "VbMsgBoxResult", "MsgBoxResult");

  const string TokenBreak = "[ ,)]";
  S = RegExReplace(S, "InventFolder(" + TokenBreak + ")", "InventFolder()$1");
  S = RegExReplace(S, "PXFolder(" + TokenBreak + ")", "InventFolder()$1");
  S = RegExReplace(S, "FXFolder(" + TokenBreak + ")", "InventFolder()$1");
  S = RegExReplace(S, "InventFolder(" + TokenBreak + ")", "InventFolder()$1");
  S = RegExReplace(S, "IsDevelopment(" + TokenBreak + ")", "IsDevelopment()$1");

  ProjectSpecificPostCodeLineConvert = S;
  return ProjectSpecificPostCodeLineConvert;
}
}
