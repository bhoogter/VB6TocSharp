using System;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modUtils;


static class modConfig
{
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


    public static string vbpFile
    {
        get
        {
            string vbpFile;
            LoadSettings();
            if (mVBPFile == "")
            {
                mVBPFile = def_vbpFile;
            }
            vbpFile = mVBPFile;

            return vbpFile;
        }
    }
    public static string vbpPath
    {
        get
        {
            string vbpPath;
            vbpPath = FilePath(vbpFile);

            return vbpPath;
        }
    }


    public static string INIFile()
    {
        string INIFile = "";
        INIFile = AppDomain.CurrentDomain.BaseDirectory + "\\VB6toCS.INI";
        return INIFile;
    }

    public static void LoadSettings(bool Force = false)
    {
        if (Loaded && !Force)
        {
            return;

        }
        Loaded = true;
        mVBPFile = modINI.INIRead(INISection_Settings, INIKey_VBPFile, INIFile());
        mOutputFolder = modINI.INIRead(INISection_Settings, INIKey_OutputFolder, INIFile());
        mAssemblyName = modINI.INIRead(INISection_Settings, INIKey_AssemblyName, INIFile());
    }

    public static string OutputFolder(string F = "")
    {
        string OutputFolder = "";
        LoadSettings();
        if (mOutputFolder == "")
        {
            mOutputFolder = def_outputFolder;
        }
        OutputFolder = mOutputFolder;
        if (Right(OutputFolder, 1) != "\\")
        {
            OutputFolder = OutputFolder + "\\";
        }
        OutputFolder = OutputFolder + OutputSubFolder(F);
        if (Dir(OutputFolder, vbDirectory) == "")
        {
            // TODO (not supported): On Error GoTo CantMakeOutputFolder
            MkDir(OutputFolder);
        }
        return OutputFolder;

    CantMakeOutputFolder:;
        if (!Hush)
        {
            MsgBox("Failed creating folder.  Perhaps create it yourself?" + vbCrLf + OutputFolder);
        }
        return OutputFolder;
    }

    public static string AssemblyName()
    {
        string AssemblyName = "";
        LoadSettings();
        if (mAssemblyName == "")
        {
            mAssemblyName = def_AssemblyName;
        }
        AssemblyName = mAssemblyName;
        return AssemblyName;
    }

    public static string OutputSubFolder(string F)
    {
        string OutputSubFolder = "";
        LoadSettings();
        switch (FileExt(F))
        {
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
