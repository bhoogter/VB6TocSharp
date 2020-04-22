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


namespace VB2CS.Forms
{
public partial class frmConfig : Window {
  private static frmConfig _instance;
  public static frmConfig instance { set { _instance = null; } get { return _instance ?? (_instance = new frmConfig()); }}  public static void Load() { if (instance == null) { dynamic A = frmConfig.instance; } }  public static void Unload() { if (instance != null) instance.Close(); instance = null; }  public frmConfig() { InitializeComponent(); }


// Option Explicit //Right Justify


private void Form_Load(object sender, RoutedEventArgs e) { Form_Load(); }
private void Form_Load() {
  txtVBPFile.Text = modConfig.vbpFile;
  txtOutput.Text = modConfig.OutputFolder;
  txtAssemblyName.Text = modConfig.AssemblyName;
}

private void cmdCancel_Click(object sender, RoutedEventArgs e) { cmdCancel_Click(); }
private void cmdCancel_Click() {
  Unload(this);
}

private void cmdOK_Click(object sender, RoutedEventArgs e) { cmdOK_Click(); }
private void cmdOK_Click() {
  modINI.INIWrite(INISection_Settings, INIKey_VBPFile, txtVBPFile.Text, INIFile());
  modINI.INIWrite(INISection_Settings, INIKey_OutputFolder, txtOutput.Text, INIFile());
  modINI.INIWrite(INISection_Settings, INIKey_AssemblyName, txtAssemblyName.Text, INIFile());
  modConfig.LoadSettings(true);
  Unload(this);
}

private void txtOutput_Validate(ref bool Cancel_UNUSED) {
  if (Dir(txtOutput.Text, vbDirectory) == "") {
    MsgBox("Output folder does not exist.  Please create to prevent errors.");
  }
}

private void txtVBPFile_Validate(ref bool Cancel_UNUSED) {
  if (Dir(txtVBPFile.Text) == "") {
    MsgBox("Project file does not exist.  Please give a valid project to prevent errors.");
  }
}

private void txtAssemblyName_Validate(ref bool Cancel_UNUSED) {
  if (txtAssemblyName.Text == "") {
    MsgBox("Please enter something for an assembly name.");
  }
}

private void txtVBPFile_GotFocus(object sender, RoutedEventArgs e) { txtVBPFile_GotFocus(); }
private void txtVBPFile_GotFocus() {
  txtVBPFile.SelectionStart = 0;
  txtVBPFile.SelectionLength = Len(txtVBPFile.Text);
}

private void txtOutput_GotFocus(object sender, RoutedEventArgs e) { txtOutput_GotFocus(); }
private void txtOutput_GotFocus() {
  txtOutput.SelectionStart = 0;
  txtOutput.SelectionLength = Len(txtOutput.Text);
}

private void txtAssemblyName_GotFocus(object sender, RoutedEventArgs e) { txtAssemblyName_GotFocus(); }
private void txtAssemblyName_GotFocus() {
  txtAssemblyName.SelectionStart = 0;
  txtAssemblyName.SelectionLength = Len(txtAssemblyName.Text);
}


}
}
