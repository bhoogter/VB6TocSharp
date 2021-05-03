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


static class modProjectFiles {
// Option Explicit


public static string VBPModules(string ProjectFile= "") {
  string VBPModules = "";
  string S = "";
  dynamic L = null;

  string T = "";

  const string C = "Module=";
  if (ProjectFile == "") {
    ProjectFile = vbpFile;
  }
  S = ReadEntireFile(ProjectFile);
  foreach(var iterL in Split(S, vbCrLf)) {
L = iterL;
    if (Left(L, Len(C)) == C) {
      T = Mid(L, Len(C) + 1);
      if (IsInStr(T, ";")) {
        T = SplitWord(T, 2, ";");
      }
//If IsInStr(LCase(T), "subclass") Then Stop
      if (LCase(T) == "modlistsubclass.bas") {
goto NextItem;
      }
      VBPModules = VBPModules + IIf(VBPModules == "", "", vbCrLf) + T;
    }
NextItem:;
  }
  return VBPModules;
}

public static string VBPForms(string ProjectFile= "") {
  string VBPForms = "";
  const bool WithExt = true;
  string S = "";
  dynamic L = null;

  string T = "";

  const string C = "Form=";
  if (ProjectFile == "") {
    ProjectFile = vbpFile;
  }
  S = ReadEntireFile(ProjectFile);
  foreach(var iterL in Split(S, vbCrLf)) {
L = iterL;
    if (Left(L, Len(C)) == C) {
      T = Mid(L, Len(C) + 1);
      if (IsInStr(T, ";")) {
        T = SplitWord(T, 1, ";");
      }
      if (!WithExt && Right(T, 4) == ".frm") {
        T = Left(T, Len(T) - 4);
      }
      switch(LCase(T)) {
        case "faxtest":
          T = "FaxPO";
          break;
        case "frmpos":
          T = "frmCashRegister";
          break;
        case "frmposquantity":
          T = "frmCashRegisterQuantity";
          break;
        case "calendarinst":
          T = "CalendarInstr";
          break;
        case "frmedi":
          T = "frmAshleyEDIItemAlign";
          break;
        case "frmpracticefiles":
          T = "PracticeFiles";
          break;
        case "txttextselect":
          T = "frmSelectText";
break;
}
      VBPForms = VBPForms + IIf(VBPForms == "", "", vbCrLf) + T;
    }
NextItem:;
  }
  return VBPForms;
}

public static string VBPClasses(string ProjectFile= "", bool ClassNames= false) {
  string VBPClasses = "";
  string S = "";
  dynamic L = null;

  string T = "";

  const string C = "Class=";
  if (ProjectFile == "") {
    ProjectFile = vbpFile;
  }
  S = ReadEntireFile(ProjectFile);
  foreach(var iterL in Split(S, vbCrLf)) {
L = iterL;
    if (Left(L, Len(C)) == C) {
      T = Mid(L, Len(C) + 1);
      if (IsInStr(T, ";")) {
        T = SplitWord(T, 2, ";");
      }
      VBPClasses = VBPClasses + IIf(VBPClasses == "", "", vbCrLf) + T;
    }
NextItem:;
  }
  if (ClassNames) {
    VBPClasses = Replace(VBPClasses, ".cls", "");
  }
  return VBPClasses;
}

public static string VBPUserControls(string ProjectFile= "") {
  string VBPUserControls = "";
  string S = "";
  dynamic L = null;

  string T = "";

  const string C = "UserControl=";
  if (ProjectFile == "") {
    ProjectFile = vbpFile;
  }
  S = ReadEntireFile(ProjectFile);
  foreach(var iterL in Split(S, vbCrLf)) {
L = iterL;
    if (Left(L, Len(C)) == C) {
      T = Mid(L, Len(C) + 1);
      if (IsInStr(T, ";")) {
        T = SplitWord(T, 2, ";");
      }
      VBPUserControls = VBPUserControls + IIf(VBPUserControls == "", "", vbCrLf) + T;
    }
NextItem:;
  }
  return VBPUserControls;
}
}
