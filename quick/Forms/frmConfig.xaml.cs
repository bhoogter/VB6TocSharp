using System.Windows;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modConfig;


namespace VB2CS.Forms
{
    public partial class frmConfig : Window
    {
        private static frmConfig _instance;
        public static frmConfig instance { set { _instance = null; } get { return _instance ?? (_instance = new frmConfig()); } }
        public static void Load() { if (_instance == null) { dynamic A = frmConfig.instance; } }
        public static void Unload() { if (_instance != null) instance.Close(); _instance = null; }
        public frmConfig() { InitializeComponent(); }


        private void Form_Load(object sender, RoutedEventArgs e) { Form_Load(); }
        private void Form_Load()
        {
            modConfig.Hush = true;
            txtVBPFile.Text = modConfig.vbpFile;
            txtOutput.Text = modConfig.OutputFolder();
            txtAssemblyName.Text = modConfig.AssemblyName();
            modConfig.Hush = false;
        }
        private void cmdCancel_Click(object sender, RoutedEventArgs e) { cmdCancel_Click(); }
        private void cmdCancel_Click()
        {
            Unload();
        }
        private void cmdOK_Click(object sender, RoutedEventArgs e) { cmdOK_Click(); }
        private void cmdOK_Click()
        {
            modINI.INIWrite(INISection_Settings, INIKey_VBPFile, txtVBPFile.Text, INIFile);
            modINI.INIWrite(INISection_Settings, INIKey_OutputFolder, txtOutput.Text, INIFile);
            modINI.INIWrite(INISection_Settings, INIKey_AssemblyName, txtAssemblyName.Text, INIFile);
            modConfig.LoadSettings(true);
            Unload();
        }
        private void txtOutput_Validate(ref bool Cancel)
        {
            if (Dir(txtOutput.Text, vbDirectory) == "")
            {
                MsgBox("Output folder does not exist.  Please create to prevent errors.");
            }
        }
        private void txtVBPFile_Validate(ref bool Cancel)
        {
            if (Dir(txtVBPFile.Text) == "")
            {
                MsgBox("Project file does not exist.  Please give a valid project to prevent errors.");
            }
        }
        private void txtAssemblyName_Validate(ref bool Cancel)
        {
            if (txtAssemblyName.Text == "")
            {
                MsgBox("Please enter something for an assembly name.");
            }
        }
        private void txtVBPFile_GotFocus(object sender, RoutedEventArgs e) { txtVBPFile_GotFocus(); }
        private void txtVBPFile_GotFocus() { txtVBPFile.SelectionStart = 0; txtVBPFile.SelectionLength = Len(txtVBPFile.Text); }
        private void txtOutput_GotFocus(object sender, RoutedEventArgs e) { txtOutput_GotFocus(); }
        private void txtOutput_GotFocus() { txtOutput.SelectionStart = 0; txtOutput.SelectionLength = Len(txtOutput.Text); }
        private void txtAssemblyName_GotFocus(object sender, RoutedEventArgs e) { txtAssemblyName_GotFocus(); }
        private void txtAssemblyName_GotFocus() { txtAssemblyName.SelectionStart = 0; txtAssemblyName.SelectionLength = Len(txtAssemblyName.Text); }

    }
}
