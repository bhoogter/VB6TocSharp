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


static class modDirStack {
// Option Explicit
private static Collection DirStack = new Collection();


public static string PushDir(string NewDir, bool doSet= true) {
  string PushDir = "";
//::::PushDir
//:::SUMMARY
//:Basic Directory Stack - Push cur dir to stack and CD to parameter.
//:::DESCRIPTION
//:1. Push Current Dir to stack
//:2. CD to new folder.
//:::PAREMETERS
//: - sNewDir - String - Directory to CD into.
//: - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
//:::RETURNS
//:Returns current directory.
//:::SEE ALSO
//: PopDir, PeekDir
  int N = 0;


  // TODO (not supported): On Error Resume Next
  if (DirStack == null) {
    DirStack = new Collection();;
    DirStack.Add(0, "n");
  }

  N = Val(DirStack.Item("n")) + 1;
  DirStack.Remove("n");
  DirStack.Add(N, "n");
  DirStack.Add(CurDir, "_" + N);

  if (doSet) {
    ChDir(NewDir);
  }

  PushDir = CurDir;
  return PushDir;
}

public static string PopDir(bool doSet= true) {
  string PopDir = "";
//::::PopDir
//:::SUMMARY
//:Remove to dir from stack.  Error Safe.  Generally to change current directory.
//:::DESCRIPTION
//:1. Pop Dir from stack.
//:2. CD to dir.
//:::PAREMETERS
//: - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
//:::RETURNS
//:Returns directory popped.
//:::SEE ALSO
//: PopDir, PeekDir
  int N = 0;
  string V = "";


  // TODO (not supported): On Error Resume Next
  if (DirStack == null) {
    return PopDir;

  }

  N = Val(DirStack.Item("n"));
  PopDir = DirStack.Item("_" + N);

  if (N > 1) {
    N = N - 1;
    DirStack.Remove("n");
    DirStack.Add(N, "n");
  } else {
    DirStack = null;
  }

  if (doSet) {
    ChDir(PopDir);
  }
  return PopDir;
}

public static string PeekDir(bool doSet= true) {
  string PeekDir = "";
//::::PeekDir
//:::SUMMARY
//:Return directory on top of stack without removing it.  Generally to change current directory.
//:::DESCRIPTION
//:1. Push Current Dir to stack
//:2. CD to new folder.
//:::PAREMETERS
//: - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
//:::RETURNS
//:Returns top stack item (without removing it from stack).
//:::SEE ALSO
//: PopDir, PeekDir
  int N = 0;
  string V = "";


  // TODO (not supported): On Error Resume Next
  if (DirStack == null) {
    return PeekDir;

  }

  N = Val(DirStack.Item("n"));
  PeekDir = DirStack.Item("_" + N);

  if (doSet) {
    ChDir(PeekDir);
  }
  return PeekDir;
}
}
