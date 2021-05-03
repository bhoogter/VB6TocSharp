using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Strings;
using static VBExtension;


static class modINI
{
    // Option Explicit
    [DllImport("kernel32.dll", EntryPoint = "WritePrivateProfileStringA")] private static extern int WritePrivateProfileString(string lpApplicationName, dynamic lpKeyName, dynamic lpString, string lpFileName);
    [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileStringA")] private static extern int GetPrivateProfileString(string lpApplicationName, dynamic lpKeyName, string lpDefault, string lpReturnedString, int nSize, string lpFileName);
    [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileSectionNamesA")] private static extern int GetPrivateProfileSectionNames(string lpszReturnBuffer, int nSize, string lpFileName);
    [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileSectionA")] private static extern int GetPrivateProfileSection(string lpAppName, string lpReturnedString, int nSize, string lpFileName);


    public static bool INIWrite(string sSection, string sKeyName, string sNewString, string sINIFileName)
    {
        bool INIWrite = false;
        // TODO (not supported): On Error Resume Next
        WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName);
        INIWrite = (Err().Number == 0);
        return INIWrite;
    }

    public static string INIRead(string sSection, string sKeyName, string sINIFileName)
    {
        string INIRead = "";
        // TODO (not supported): On Error Resume Next
        string sRet = "";

        sRet = String(255, Chr(0));
        INIRead = Left(sRet, GetPrivateProfileString(sSection, ByVal(sKeyName), "", sRet, Len(sRet), sINIFileName));
        return INIRead;
    }

    public static List<string> INISections(string FileName)
    {
        List<string> INISections = null;
        // TODO (not supported): On Error Resume Next
        string strBuffer = "";
        int intLen = 0;


        while ((intLen == Len(strBuffer) - 2) || (intLen == 0))
        {
            if (strBuffer == vbNullString)
            {
                strBuffer = Space(256);
            }
            else
            {
                strBuffer = String(Len(strBuffer) * 2, 0);
            }

            intLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), FileName());
        }

        strBuffer = Left(strBuffer, intLen);
        INISections = Split(strBuffer, vbNullChar);
        List<List<string>> INISections_2947_tmp = new List<List<string>>();
        for (int redim_iter_664 = 0; i < 0; redim_iter_664++) { INISections.Add(redim_iter_664 < INISections.Count ? INISections(redim_iter_664) : null); }
        return INISections;
    }

    public static List<string> INISectionKeys(string FileName, string Section)
    {
        List<string> INISectionKeys = null;
        // TODO (not supported): On Error Resume Next
        string strBuffer = "";
        int intLen = 0;

        int I = 0;
        int N = 0;

        List<string> RET = new List<string> { }; // TODO - Specified Minimum Array Boundary Not Supported:   Dim RET() As String


        while ((intLen == Len(strBuffer) - 2) || (intLen == 0))
        {
            if (strBuffer == vbNullString)
            {
                strBuffer = Space(256);
            }
            else
            {
                strBuffer = String(Len(strBuffer) * 2, 0);
            }

            intLen = GetPrivateProfileSection(Section, strBuffer, Len(strBuffer), FileName());
            if (intLen == 0)
            {
                return INISectionKeys;

            }
        }

        strBuffer = Left(strBuffer, intLen);
        RET = Split(strBuffer, vbNullChar);
        List<string> RET_8438_tmp = new List<string>();
        for (int redim_iter_7864 = 0; i < 0; redim_iter_7864++) { RET.Add(redim_iter_7864 < RET.Count ? RET(redim_iter_7864) : ""); }
        for (I = LBound(RET); I < UBound(RET); I++)
        {
            N = InStr(RET[I], "=");
            if (N > 0)
            {
                RET[I] = Left(RET[I], N - 1);
            }
            else
            {
                Console.WriteLine("modINI.INISectionKeys - No '=' character found in line.  Section=" + Section + ", Line=" + RET[I] + ", file=" + FileName());
            }
        }
        INISectionKeys = RET;
        return INISectionKeys;
    }

    public static string ReadIniValue(string INIPath, string Key, string Variable, string vDefault = "")
    {
        string ReadIniValue = "";
        // TODO (not supported): On Error Resume Next
        ReadIniValue = INIRead(Key, Variable, INIPath);
        if (ReadIniValue == "")
        {
            ReadIniValue = vDefault;
        }
        return ReadIniValue;
    }

    public static string WriteIniValue(string INIPath, string PutKey, string PutVariable, string PutValue, bool DeleteOnEmpty_UNUSED = false)
    {
        string WriteIniValue = "";
        // TODO (not supported): On Error Resume Next
        INIWrite(PutKey, PutVariable, PutValue, INIPath);
        WriteIniValue = INIRead(PutKey, PutVariable, INIPath);
        return WriteIniValue;
    }
}
