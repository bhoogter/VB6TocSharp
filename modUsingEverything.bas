Attribute VB_Name = "modUsingEverything"
Option Explicit

' Provide a mechanism to add a using for everything pertinent to every CS file.

Private Everything As String
Private Const VB6Compat As String = "Microsoft.VisualBasic.Compatibility.VB6"


' Returns preamble for every CS file generated.
Public Function UsingEverything(Optional ByVal PackageName As String = "") As String
  Dim List As String, Path As String, Name As String
  Dim E As String, L As Variant
  Dim R As String, N As String, M As String
  E = ""
  R = "": N = vbCrLf: M = ""
  
  If PackageName <> "" Then
'    R = R & N & "package " & PackagePrefix & PackageName & ";"
    R = R & N & ""
  End If
  
  If Everything = "" Then
    E = E & M & "using VB6 = " & VB6Compat & ";"
    E = E & N & "using System.Runtime.InteropServices;"
    E = E & N & "using static VBExtension;"
    E = E & N & "using static VBConstants;"
    E = E & N & "using Microsoft.VisualBasic;"
    
    E = E & N & "using System;"
    E = E & N & "using System.Windows;"
    E = E & N & "using System.Windows.Controls;"
    E = E & N & "using static System.DateTime;"
    E = E & N & "using static System.Math;"
    E = E & N & "using System.Linq;"
    
    E = E & N & "using static Microsoft.VisualBasic.Globals;"
    E = E & N & "using static Microsoft.VisualBasic.Collection;"
    E = E & N & "using static Microsoft.VisualBasic.Constants;"
    E = E & N & "using static Microsoft.VisualBasic.Conversion;"
    E = E & N & "using static Microsoft.VisualBasic.DateAndTime;"
    E = E & N & "using static Microsoft.VisualBasic.ErrObject;"
    E = E & N & "using static Microsoft.VisualBasic.FileSystem;"
    E = E & N & "using static Microsoft.VisualBasic.Financial;"
    E = E & N & "using static Microsoft.VisualBasic.Information;"
    E = E & N & "using static Microsoft.VisualBasic.Interaction;"
    E = E & N & "using static Microsoft.VisualBasic.Strings;"
    E = E & N & "using static Microsoft.VisualBasic.VBMath;"
    E = E & N & "using System.Collections.Generic;"
    
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ColorConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.DrawStyleConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.FillStyleConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.GlobalModule;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.Printer;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterCollection;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterObjectConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ScaleModeConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.SystemColorConstants;"
    E = E & N & "using ADODB;"
    
    E = E & N & "using System;"
    E = E & N & "using System.Collections.Generic;"
    E = E & N & "using System.Linq;"
    E = E & N & "using System.Text;"
    E = E & N & "using System.Threading.Tasks;"
    E = E & N & "using System.Windows;"
    E = E & N & "using System.Windows.Controls;"
    E = E & N & "using System.Windows.Data;"
    E = E & N & "using System.Windows.Documents;"
    E = E & N & "using System.Windows.Input;"
    E = E & N & "using System.Windows.Media;"
    E = E & N & "using System.Windows.Media.Imaging;"
    E = E & N & "using System.Windows.Shapes;"
    
    E = E & N
    
    E = E & N & "using " & AssemblyName & ".Forms;"
    
    Path = FilePath(vbpFile)
    For Each L In Split(VBPModules(vbpFile), vbCrLf)
      If L <> "" Then
        Name = ModuleName(ReadEntireFile(Path & L))
        E = E & N & "using static " & PackagePrefix & Name & ";"
      End If
    Next
    For Each L In Split(VBPForms(vbpFile), vbCrLf)
      If L <> "" Then
        Name = ModuleName(ReadEntireFile(Path & L))
        E = E & N & "using static " & AssemblyName & ".Forms." & Name & ";"
      End If
    Next
    For Each L In Split(VBPClasses(vbpFile), vbCrLf)
      If L <> "" Then
        Name = ModuleName(ReadEntireFile(Path & L))
        E = E & N & "using static " & AssemblyName & ".Classes." & Name & ";"
      End If
    Next
    Everything = E
  End If
  
  R = Everything & N & R
  UsingEverything = R
End Function
