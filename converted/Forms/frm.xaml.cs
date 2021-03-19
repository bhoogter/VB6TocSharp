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


namespace VB2CS.Forms
{
public partial class frm : Window {
  private static frm _instance;
  public static frm instance { set { _instance = null; } get { return _instance ?? (_instance = new frm()); }}  public static void Load() { if (_instance == null) { dynamic A = frm.instance; } }  public static void Unload() { if (_instance != null) instance.Close(); _instance = null; }  public frm() { InitializeComponent(); }


// Option Explicit //Right Justify
public int pMax = 0;


private void cmdAll_Click(object sender, RoutedEventArgs e) { cmdAll_Click(); }
private void cmdAll_Click() {
  if (!ConfigValid) {
return;

  }
  IsWorking();
  ConvertProject(txtSrc.Text);
  IsWorking(true);
}

private void cmdClasses_Click(object sender, RoutedEventArgs e) { cmdClasses_Click(); }
private void cmdClasses_Click() {
  if (!ConfigValid) {
return;

  }
  IsWorking();
  ConvertFileList(FilePath(txtSrc.Text), VBPClasses(txtSrc.Text));
  IsWorking(true);
}

private void cmdConfig_Click(object sender, RoutedEventArgs e) { cmdConfig_Click(); }
private void cmdConfig_Click() {
  frmConfig.Show(1);
  modConfig.LoadSettings();
}

private void cmdExit_Click(object sender, RoutedEventArgs e) { cmdExit_Click(); }
private void cmdExit_Click() {
  Unload(this);
}

private void cmdFile_Click(object sender, RoutedEventArgs e) { cmdFile_Click(); }
private void cmdFile_Click() {
  if (!ConfigValid) {
return;

  }
  IsWorking();
  ConvertFile(txtFile.Text);
  IsWorking(true);
  MsgBox("Converted " + txtFile.Text + ".");
}

private void cmdForms_Click(object sender, RoutedEventArgs e) { cmdForms_Click(); }
private void cmdForms_Click() {
  if (!ConfigValid) {
return;

  }
  IsWorking();
  ConvertFileList(FilePath(txtSrc.Text), VBPForms(txtSrc.Text));
  IsWorking(true);
}

private void cmdModules_Click(object sender, RoutedEventArgs e) { cmdModules_Click(); }
private void cmdModules_Click() {
  if (!ConfigValid) {
return;

  }
  IsWorking();
  ConvertFileList(FilePath(txtSrc.Text), VBPModules(txtSrc.Text));
  IsWorking(true);
}

private bool ConfigValid() {
  bool ConfigValid = false;
  modConfig.LoadSettings();

  if (Dir(modConfig.vbpFile) == "") {
    MsgBox("Project file not found.  Perhaps do config first?", vbExclamation, "File Not Found");
    return ConfigValid;

  }
  if (Dir(modConfig.OutputFolder, vbDirectory) == "") {
    MsgBox("Ouptut Folder not found.  Perhaps do config first?", vbExclamation, "Directory Not Found");
    return ConfigValid;

  }
  if (modConfig.AssemblyName == "") {
    MsgBox("Assembly name not set.  Perhaps do config first?", vbExclamation, "Setting Not Found");
    return ConfigValid;

  }
  ConfigValid = true;
  return ConfigValid;
}

private void IsWorking(bool Done= false) {
  txtFile.IsEnabled = Done;
  cmdConfig.IsEnabled = Done;
  cmdLint.IsEnabled = Done;
  cmdFile.IsEnabled = Done;
  cmdAll.IsEnabled = Done;
  cmdClasses.IsEnabled = Done;
  cmdExit.IsEnabled = Done;
  cmdForms.IsEnabled = Done;
  cmdModules.IsEnabled = Done;
  txtSrc.IsEnabled = Done;
  cmdScan.IsEnabled = Done;
  cmdSupport.IsEnabled = Done;
  MousePointer = IIf(Done, vbDefault, vbHourglass);
}

public dynamic Prg(int Val= -1, int Max= -1, string Cap= "#") {
  dynamic Prg = null;
  // TODO (not supported): On Error Resume Next
  if (Max >= 0) {
    pMax = Max;
  }
  lblPrg.DefaultProperty = IIf(Prg == "#", "", Cap);
  shpPrg.Width = Val / pMax * 2415;
  shpPrg.Visibility = Val >= 0;
  lblPrg.Visibility = shpPrg.Visibility;
  return Prg;
}

private void cmdLint_Click(object sender, RoutedEventArgs e) { cmdLint_Click(); }
private void cmdLint_Click() {
  if (!ConfigValid) {
return;

  }
  LintFolder();
}

private void cmdScan_Click(object sender, RoutedEventArgs e) { cmdScan_Click(); }
private void cmdScan_Click() {
  if (!ConfigValid) {
return;

  }
  IsWorking(false);
  ScanRefs();
  IsWorking(true);
}

private void cmdSupport_Click(object sender, RoutedEventArgs e) { cmdSupport_Click(); }
private void cmdSupport_Click() {
  if (!ConfigValid) {
return;

  }
  if (MsgBox("Generate Project files?", vbYesNo) == vbYes) {
    CreateProjectFile(vbpFile);
  }
  if (MsgBox("Generate Support files?", vbYesNo) == vbYes) {
    CreateProjectSupportFiles();
  }
}

private void Form_Load(object sender, RoutedEventArgs e) { Form_Load(); }
private void Form_Load() {
  modConfig.LoadSettings();
  txtSrc.Text = vbpFile;
}


}
}
