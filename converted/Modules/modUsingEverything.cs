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


static class modUsingEverything {
// Option Explicit
private static string Everything = "";
private const string VB6Compat = "Microsoft.VisualBasic.Compatibility.VB6";


public static string UsingEverything(string PackageName= "") {
  string UsingEverything = "";
  string List = "";
  string Path = "";
  string Name = "";

  string E = "";
  dynamic L = null;

  string R = "";
  string N = "";
  string M = "";

  E = "";
  R = "";
  N = vbCrLf;
  M = "";

  if (PackageName != "") {
//    R = R & N & "package " & PackagePrefix & PackageName & ";"
    R = R + N + "";
  }

  if (Everything == "") {
    E = E + M + "using VB6 = " + VB6Compat + ";";
    E = E + N + "using System.Runtime.InteropServices;";
    E = E + N + "using static VBExtension;";
    E = E + N + "using static VBConstants;";
    E = E + N + "using Microsoft.VisualBasic;";

    E = E + N + "using System;";
    E = E + N + "using System.Windows;";
    E = E + N + "using System.Windows.Controls;";
    E = E + N + "using static System.DateTime;";
    E = E + N + "using static System.Math;";

    E = E + N + "using static Microsoft.VisualBasic.Globals;";
    E = E + N + "using static Microsoft.VisualBasic.Collection;";
    E = E + N + "using static Microsoft.VisualBasic.Constants;";
    E = E + N + "using static Microsoft.VisualBasic.Conversion;";
    E = E + N + "using static Microsoft.VisualBasic.DateAndTime;";
    E = E + N + "using static Microsoft.VisualBasic.ErrObject;";
    E = E + N + "using static Microsoft.VisualBasic.FileSystem;";
    E = E + N + "using static Microsoft.VisualBasic.Financial;";
    E = E + N + "using static Microsoft.VisualBasic.Information;";
    E = E + N + "using static Microsoft.VisualBasic.Interaction;";
    E = E + N + "using static Microsoft.VisualBasic.Strings;";
    E = E + N + "using static Microsoft.VisualBasic.VBMath;";
    E = E + N + "using System.Collections.Generic;";

    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ColorConstants;";
    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.DrawStyleConstants;";
    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.FillStyleConstants;";
    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.GlobalModule;";
    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.Printer;";
    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterCollection;";
    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterObjectConstants;";
    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ScaleModeConstants;";
    E = E + N + "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.SystemColorConstants;";
    E = E + N + "using ADODB;";

    E = E + N + "using System;";
    E = E + N + "using System.Collections.Generic;";
    E = E + N + "using System.Linq;";
    E = E + N + "using System.Text;";
    E = E + N + "using System.Threading.Tasks;";
    E = E + N + "using System.Windows;";
    E = E + N + "using System.Windows.Controls;";
    E = E + N + "using System.Windows.Data;";
    E = E + N + "using System.Windows.Documents;";
    E = E + N + "using System.Windows.Input;";
    E = E + N + "using System.Windows.Media;";
    E = E + N + "using System.Windows.Media.Imaging;";
    E = E + N + "using System.Windows.Shapes;";

    E = E + N;

    E = E + N + "using " + AssemblyName() + ".Forms;";

    Path = FilePath(vbpFile);
    foreach(var iterL in Split(VBPModules(vbpFile), vbCrLf)) {
L = iterL;
      if (L != "") {
        Name = ModuleName(ReadEntireFile(Path + L));
        E = E + N + "using static " + PackagePrefix + Name + ";";
      }
    }
    foreach(var iterL in Split(VBPForms(vbpFile), vbCrLf)) {
L = iterL;
      if (L != "") {
        Name = ModuleName(ReadEntireFile(Path + L));
        E = E + N + "using static " + AssemblyName() + ".Forms." + Name + ";";
      }
    }
//    For Each L In Split(VBPClasses(vbpFile), vbCrLf)  ' controls?
//      If L <> "" Then
//        Name = ModuleName(ReadEntireFile(Path & L))
//        E = E & N & "using " & PackagePrefix & Name & ";"
//      End If
//    Next
    Everything = E;
  }

  R = Everything + N + R;
  UsingEverything = R;
  return UsingEverything;
}
}
