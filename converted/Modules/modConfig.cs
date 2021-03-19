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


static class modConfig {
// Option Explicit
public const int SpIndent = 2;
public const string DefaultDataType = "dynamic";
public const string PackagePrefix = "";
private const string def_vbpFile = "C:\\WinCDS.NET\\cnv\\prj.vbp";
private const string def_outputFolder = "C:\\WinCDS.NET\\cnv\\converted\\";
private const string def_AssemblyName = "VB2CS";
private static string mVBPFile = "";
private static string mOutputFolder = "";
private static string mAssemblyName = "";
private static bool Loaded = false;
public static bool Hush = false;
public const string INISection_Settings = "Settings";
public const string INIKey_VBPFile = "VBPFile";
public const string INIKey_OutputFolder = "OutputFolder";
public const string INIKey_AssemblyName = "AssemblyName";


public static string vbpFile {
  get {
    string vbpFile;
    LoadSettings();
if (mVBPFile == "") {
  mVBPFile = def_vbpFile;
}
vbpFile = mVBPFile;

  return vbpFile;
  }
}
public static string vbpPath {
  get {
    string vbpPath;
    vbpPath = FilePath(vbpFile);

  return vbpPath;
  }
}


public static string INIFile() {
  string INIFile = "";
  INIFile = App.Path + "\\VB6toCS.INI";
  return INIFile;
}

public static void LoadSettings(bool Force= false) {
  if (Loaded && !Force) {
return;

  }
  Loaded = true;
  mVBPFile = modINI.INIRead(INISection_Settings, INIKey_VBPFile, INIFile());
  mOutputFolder = modINI.INIRead(INISection_Settings, INIKey_OutputFolder, INIFile());
  mAssemblyName = modINI.INIRead(INISection_Settings, INIKey_AssemblyName, INIFile());
}

public static string OutputFolder(string F= "") {
  string OutputFolder = "";
  LoadSettings();
  if (mOutputFolder == "") {
    mOutputFolder = def_outputFolder;
  }
  OutputFolder = mOutputFolder;
  if (Right(OutputFolder, 1) != "\\") {
    OutputFolder = OutputFolder + "\\";
  }
  OutputFolder = OutputFolder + OutputSubFolder(F);
  if (Dir(OutputFolder, vbDirectory) == "") {
    // TODO (not supported): On Error GoTo CantMakeOutputFolder
    MkDir(OutputFolder);
  }
  return OutputFolder;

CantMakeOutputFolder:
  MsgBox("Failed creating folder.  Perhaps create it yourself?" + vbCrLf + OutputFolder);
  return OutputFolder;
}

public static string AssemblyName() {
  string AssemblyName = "";
  LoadSettings();
  if (mAssemblyName == "") {
    mAssemblyName = def_AssemblyName;
  }
  AssemblyName = mAssemblyName;
  return AssemblyName;
}

public static string OutputSubFolder(string F) {
  string OutputSubFolder = "";
  LoadSettings();
  switch(FileExt(F)) {
    case ".bas":
      OutputSubFolder = "Modules\\";
      break;
    case ".cls":
      OutputSubFolder = "Classes\\";
      break;
    case ".frm":
      OutputSubFolder = "Forms\\";
      break;
    default:
      OutputSubFolder = "";
break;
}
  return OutputSubFolder;
}
}
