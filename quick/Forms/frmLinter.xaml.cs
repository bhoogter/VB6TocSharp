using System.Windows;
using static Microsoft.VisualBasic.Strings;


namespace VB2CS.Forms
{
    public partial class frmLinter : Window
    {
        private static frmLinter _instance;
        public static frmLinter instance { set { _instance = null; } get { return _instance ?? (_instance = new frmLinter()); } }
        public static void Load() { if (_instance == null) { dynamic A = frmLinter.instance; } }
        public static void Unload() { if (_instance != null) instance.Close(); _instance = null; }
        public frmLinter() => InitializeComponent();


        // Linting Form
        private void Form_Load(object sender, RoutedEventArgs e) { Form_Load(); }
        private void Form_Load()
        {
            txtVBPFile.Text = modConfig.vbpFile;
            txtFile.Text = "";
        }
        private void cmdClose_Click(object sender, RoutedEventArgs e) { cmdClose_Click(); }
        private void cmdClose_Click()
        {
            Unload();
        }
        private void cmdLint_Click(object sender, RoutedEventArgs e) { cmdLint_Click(); }
        private void cmdLint_Click()
        {
            string File = "";
            string Results = "";
            fraConfig.IsEnabled = false;
            if (txtFile.Text == "")
            {
                Results = modQuickLint.Lint();
            }
            else
            {
                File = txtFile.Text;
                if (InStr(File, "\\") == 0) File = Left(txtVBPFile.Text, InStrRev(txtVBPFile.Text, "\\")) + File;
                Results = modQuickLint.Lint(File);
            }
            fraConfig.IsEnabled = true;
            txtResults.Text = (Results == "" ? "Done." : Results);
        }

    }
}
