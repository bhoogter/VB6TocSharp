using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Interaction;
using static modConfig;
using static modConvert;
using static modProjectFiles;
using static modRefScan;
using static modSupportFiles;
using static modUtils;
using static VBConstants;
using static VBExtension;


namespace VB2CS.Forms
{
    public partial class frm : Window
    {
        private static frm _instance;
        public static frm instance { set { _instance = null; } get { return _instance ?? (_instance = new frm()); } }
        public static void Load() { if (_instance == null) { dynamic A = frm.instance; } }
        public static void Unload() { if (_instance != null) instance.Close(); _instance = null; }
        public frm() { InitializeComponent(); }


        public List<RadioButton> optVersion { get => VBExtension.controlArray<RadioButton>(this, "optVersion"); }

        // Main form
        public int pMax = 0;
        public string ConverterVersion
        {
            get
            {
                for (int I = 0; I < optVersion.Count; I += 1)
                    if (optVersion[I].IsChecked == true) return optVersion[I].Content.ToString();
                return CONVERTER_VERSION_1;
            }
        }

        private void cmdAll_Click(object sender, RoutedEventArgs e) { cmdAll_Click(); }
        private void cmdAll_Click()
        {
            if (!ConfigValid()) return;
            IsWorking();
            ConvertProject(txtSrc.Text, ConverterVersion);
            IsWorking(true);
            MsgBox("Complete");
        }
        private void cmdClasses_Click(object sender, RoutedEventArgs e) { cmdClasses_Click(); }
        private void cmdClasses_Click()
        {
            if (!ConfigValid()) return;
            IsWorking();
            ConvertFileList(FilePath(txtSrc.Text), VBPClasses(txtSrc.Text), vbCrLf, ConverterVersion);
            IsWorking(true);
        }
        private void cmdConfig_Click(object sender, RoutedEventArgs e) { cmdConfig_Click(); }
        private void cmdConfig_Click()
        {
            frmConfig.instance.ShowDialog();
            modConfig.LoadSettings();
        }
        private void cmdExit_Click(object sender, RoutedEventArgs e) { cmdExit_Click(); }
        private void cmdExit_Click()
        {
            Unload();
        }
        private void cmdFile_Click(object sender, RoutedEventArgs e) { cmdFile_Click(); }
        private void cmdFile_Click()
        {
            bool Success = false;
            if (txtFile.Text == "")
            {
                MsgBox("Enter a file in the box.", vbExclamation, "No File Entered");
                return;
            }
            if (!ConfigValid()) return;
            IsWorking();
            Success = ConvertFile(txtFile.Text, false, ConverterVersion);
            IsWorking(true);
            if (Success) MsgBox("Converted " + txtFile.Text + ".");
        }
        private void cmdForms_Click(object sender, RoutedEventArgs e) { cmdForms_Click(); }
        private void cmdForms_Click()
        {
            if (!ConfigValid()) return;
            IsWorking();
            ConvertFileList(FilePath(txtSrc.Text), VBPForms(txtSrc.Text), vbCrLf, ConverterVersion);
            IsWorking(true);
        }
        private void cmdModules_Click(object sender, RoutedEventArgs e) { cmdModules_Click(); }
        private void cmdModules_Click()
        {
            if (!ConfigValid()) return;
            IsWorking();
            ConvertFileList(FilePath(txtSrc.Text), VBPModules(txtSrc.Text), vbCrLf, ConverterVersion);
            IsWorking(true);
        }
        private bool ConfigValid()
        {
            bool _ConfigValid = false;
            modConfig.LoadSettings();
            if (Dir(modConfig.vbpFile) == "")
            {
                MsgBox("Project file not found.  Perhaps do config first?", vbExclamation, "File Not Found");
                return _ConfigValid;
            }
            if (Dir(modConfig.OutputFolder(), vbDirectory) == "")
            {
                MsgBox("Ouptut Folder not found.  Perhaps do config first?", vbExclamation, "Directory Not Found");
                return _ConfigValid;
            }
            if (modConfig.AssemblyName() == "")
            {
                MsgBox("Assembly name not set.  Perhaps do config first?", vbExclamation, "Setting Not Found");
                return _ConfigValid;
            }
            _ConfigValid = true;
            return _ConfigValid;
        }
        private void IsWorking(bool Done = false)
        {
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
            MousePointer = (Done ? vbDefault : vbHourglass);
        }
        public string Prg(int Val = -1, int Max = -1, string Cap = "#")
        {
            string _Prg = "";
            // TODO: (NOT SUPPORTED): On Error Resume Next
            if (Max >= 0) pMax = Max;
            lblPrg.Content = (_Prg == "#" ? "" : Cap);
            shpPrg.Width = Val / pMax * 2415;
            shpPrg.Visibility = Val >= 0;
            lblPrg.Visibility = shpPrg.Visibility;
            return _Prg;
        }
        private void cmdLint_Click(object sender, RoutedEventArgs e) { cmdLint_Click(); }
        private void cmdLint_Click()
        {
            if (!ConfigValid()) return;
            frmLinter.instance.ShowDialog();
        }
        private void cmdScan_Click(object sender, RoutedEventArgs e) { cmdScan_Click(); }
        private void cmdScan_Click()
        {
            if (!ConfigValid()) return;
            IsWorking(false);
            ScanRefs();
            IsWorking(true);
        }
        private void cmdSupport_Click(object sender, RoutedEventArgs e) { cmdSupport_Click(); }
        private void cmdSupport_Click()
        {
            if (!ConfigValid()) return;
            if (MsgBox("Generate Project files?", vbYesNo) == vbYes) CreateProjectFile(vbpFile);
            if (MsgBox("Generate Support files?", vbYesNo) == vbYes) CreateProjectSupportFiles();
        }
        private void Form_Load(object sender, RoutedEventArgs e) { Form_Load(); }
        private void Form_Load()
        {
            modConfig.Hush = true;
            modConfig.LoadSettings();
            modConfig.Hush = false;
            txtSrc.Text = vbpFile;
        }

    }
}
