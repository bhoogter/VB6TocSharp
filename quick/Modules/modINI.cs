using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Strings;



static class modINI
{
    [DllImport("kernel32")]
    private static extern int WritePrivateProfileString(string lpApplicationName, string lpKeyName, string lpString, string lpFileName);
    [DllImport("kernel32")]
    private static extern int GetPrivateProfileString(string lpApplicationName, string lpKeyName, string lpDefault, string lpReturnedString, int nSize, string lpFileName);
    [DllImport("kernel32.dll")]
    private static extern int GetPrivateProfileSectionNames(string lpszReturnBuffer, int nSize, string lpFileName);
    [DllImport("kernel32")]
    private static extern int GetPrivateProfileSection(string lpAppName, string lpReturnedString, int nSize, string lpFileName);
    public static bool INIWrite(string sSection, string sKeyName, string sNewString, string sINIFileName)
    {
        bool _INIWrite = false;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName);
        _INIWrite = (Err().Number == 0);
        return _INIWrite;
    }
    public static string INIRead(string sSection, string sKeyName, string sINIFileName)
    {
        string _INIRead = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        string sRet = "";
        sRet = String(255, Chr(0));
        _INIRead = Left(sRet, GetPrivateProfileString(sSection, sKeyName, "", sRet, Len(sRet), sINIFileName));
        return _INIRead;
    }
    public static List<string> INISections(string tFileName)
    {
        List<string> _INISections = null;
        // TODO: (NOT SUPPORTED): On Error Resume Next
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
            intLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), tFileName);
        }
        strBuffer = Left(strBuffer, intLen);
        _INISections = new List<string>(Split(strBuffer, vbNullChar));
        // TODO: (NOT SUPPORTED): ReDim Preserve INISections(UBound(INISections) - 1) As String
        return _INISections;
    }
    public static List<string> INISectionKeys(string tFileName, string Section)
    {
        List<string> _INISectionKeys = null;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        string strBuffer = "";
        int intLen = 0;
        int I = 0;
        int N = 0;
        List<string> Ret = new List<string>();
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
            intLen = GetPrivateProfileSection(Section, strBuffer, Len(strBuffer), tFileName);
            if (intLen == 0) return _INISectionKeys;
        }
        strBuffer = Left(strBuffer, intLen);
        Ret = new List<string>(Split(strBuffer, vbNullChar));
        // TODO: (NOT SUPPORTED): ReDim Preserve Ret(UBound(Ret) - 1) As String
        for (I = 0; I <= Ret.Count; I += 1)
        {
            N = InStr(Ret[I], "=");
            if (N > 0)
            {
                Ret[I] = Left(Ret[I], N - 1);
            }
            else
            {
                Console.WriteLine("modINI.INISectionKeys - No '=' character found in line.  Section=" + Section + ", Line=" + Ret[I] + ", file=" + tFileName);
            }
        }
        _INISectionKeys = Ret;
        return _INISectionKeys;
    }
    public static string ReadIniValue(string INIPath, string Key, string Variable, string vDefault = "")
    {
        string _ReadIniValue = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        _ReadIniValue = INIRead(Key, Variable, INIPath);
        if (_ReadIniValue == "") _ReadIniValue = vDefault;
        return _ReadIniValue;
    }
    public static string WriteIniValue(string INIPath, string PutKey, string PutVariable, string PutValue, bool DeleteOnEmpty = false)
    {
        string _WriteIniValue = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        INIWrite(PutKey, PutVariable, PutValue, INIPath);
        _WriteIniValue = INIRead(PutKey, PutVariable, INIPath);
        return _WriteIniValue;
    }

}
