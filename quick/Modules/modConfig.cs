using System;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modUtils;



static class modConfig
{
    // project config.  Handles reading/writing INI file, etc, and access to those values.
    public const int SpIndent = 2;
    public const string DefaultDataType = "dynamic";
    public const string PackagePrefix = "";
    public const string def_vbpFile = "C:\\WinCDS.NET\\cnv\\prj.vbp";
    public const string def_outputFolder = "C:\\WinCDS.NET\\cnv\\converted\\";
    public const string def_AssemblyName = "VB2CS";
    public static string mVBPFile = "";
    public static string mOutputFolder = "";
    public static string mAssemblyName = "";
    public static bool Loaded = false;
    public static bool Hush = false;
    public const string INISection_Settings = "Settings";
    public const string INIKey_VBPFile = "VBPFile";
    public const string INIKey_OutputFolder = "OutputFolder";
    public const string INIKey_AssemblyName = "AssemblyName";
    public static string vbpFile
    {
        get
        {
            string _vbpFile = default(string);
            LoadSettings();
            if (mVBPFile == "") mVBPFile = def_vbpFile;
            _vbpFile = mVBPFile;
            return _vbpFile;
        }
    }

    public static string vbpPath
    {
        get
        {
            string _vbpPath = default(string);
            _vbpPath = FilePath(vbpFile);
            return _vbpPath;
        }
    }

    public static string INIFile
    {
        get
        {
            string _INIFile = default(string);
            _INIFile = AppContext.BaseDirectory + "\\VB6toCS.INI";
            return _INIFile;
        }
    }

    public static void LoadSettings(bool Force = false)
    {
        if (Loaded && !Force) return;
        Loaded = true;
        mVBPFile = modINI.INIRead(INISection_Settings, INIKey_VBPFile, INIFile);
        mOutputFolder = modINI.INIRead(INISection_Settings, INIKey_OutputFolder, INIFile);
        mAssemblyName = modINI.INIRead(INISection_Settings, INIKey_AssemblyName, INIFile);
    }
    public static string OutputFolder(string F = "")
    {
        string _OutputFolder = "";
        LoadSettings();
        if (mOutputFolder == "") mOutputFolder = def_outputFolder;
        _OutputFolder = mOutputFolder;
        if (Right(_OutputFolder, 1) != "\\") _OutputFolder = _OutputFolder + "\\";
        _OutputFolder = _OutputFolder + OutputSubFolder(F);
        if (Dir(_OutputFolder, vbDirectory) == "")
        {
            // TODO: (NOT SUPPORTED): On Error GoTo CantMakeOutputFolder
            MkDir(_OutputFolder);
        }
        return _OutputFolder;
    CantMakeOutputFolder:;
        if (!Hush) MsgBox("Failed creating folder.  Perhaps create it yourself?" + vbCrLf + _OutputFolder);
        return _OutputFolder;
    }
    public static string AssemblyName()
    {
        string _AssemblyName = "";
        LoadSettings();
        if (mAssemblyName == "") mAssemblyName = def_AssemblyName;
        _AssemblyName = mAssemblyName;
        return _AssemblyName;
    }
    public static string OutputSubFolder(string F)
    {
        string _OutputSubFolder = "";
        LoadSettings();
        switch (FileExt(F))
        {
            case ".bas":
                _OutputSubFolder = "Modules\\";
                break;
            case ".cls":
                _OutputSubFolder = "Classes\\";
                break;
            case ".frm":
                _OutputSubFolder = "Forms\\";
                break;
            default:
                _OutputSubFolder = "";
                break;
        }
        return _OutputSubFolder;
    }

}
