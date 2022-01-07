using System.Collections.Generic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Strings;
using static modConfig;
using static modProjectFiles;
using static modTextFiles;
using static modUtils;




static class modUsingEverything
{
    public static string Everything = "";
    public const string VB6Compat = "Microsoft.VisualBasic.Compatibility.VB6";
    public static string UsingEverything(string PackageName = "")
    {
        string _UsingEverything = "";
        string List = "";
        string Path = "";
        string Name = "";
        string E = "";
        dynamic L = null;
        string R = "";
        string N = "";
        string M = "";
        E = "";
        R = ""; N = vbCrLf; M = "";
        if (PackageName != "")
        {
            // R = R & N & __S1 & PackagePrefix & PackageName & __S2
            R = R + N + "";
        }
        if (Everything == "")
        {
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
            E = E + N + "using System.Linq;";
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
            foreach (var iterL in new List<string>(Split(VBPModules(vbpFile), vbCrLf)))
            {
                L = iterL;
                if (L != "")
                {
                    Name = ModuleName(ReadEntireFile(Path + L));
                    E = E + N + "using static " + PackagePrefix + Name + ";";
                }
            }
            foreach (var iterL in new List<string>(Split(VBPForms(vbpFile), vbCrLf)))
            {
                L = iterL;
                if (L != "")
                {
                    Name = ModuleName(ReadEntireFile(Path + L));
                    E = E + N + "using static " + AssemblyName() + ".Forms." + Name + ";";
                }
            }
            // For Each L In Split(VBPClasses(vbpFile), vbCrLf)  ' controls?
            // If L <> __S1 Then
            // Name = ModuleName(ReadEntireFile(Path & L))
            // E = E & N & __S1 & PackagePrefix & Name & __S2
            // End If
            // Next
            Everything = E;
        }
        R = Everything + N + R;
        _UsingEverything = R;
        return _UsingEverything;
    }

}
