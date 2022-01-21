using System.Collections.Generic;
using static Microsoft.VisualBasic.Interaction;
using static VBExtension;



static class modRegEx
{
    // Simple regex tools for linting / converting
    public static dynamic mRegEx = null;
    private static dynamic RegEx
    {
        get
        {
            dynamic _RegEx = default(dynamic);
            if (mRegEx == null) { mRegEx = CreateObject("vbscript.regexp"); mRegEx.Global = true; }
            _RegEx = mRegEx;
            return _RegEx;
        }
    }

    // Return true/false if regex pattern `Find` is found in `Src`.
    public static bool RegExTest(string Src, string Find)
    {
        bool _RegExTest = false;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        RegEx.Pattern = Find;
        _RegExTest = RegEx.test(Src);
        return _RegExTest;
    }
    // Return number of instances of pattern `Find` in `Src`.
    public static int RegExCount(string Src, string Find)
    {
        int _RegExCount = 0;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        RegEx.Pattern = Find;
        RegEx.Global = true;
        _RegExCount = RegEx.Execute(Src).Count;
        return _RegExCount;
    }
    public static int RegExNPos(string Src, string Find, int N = 0)
    {
        int _RegExNPos = 0;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        dynamic RegM = null;
        string tempStr = "";
        string tempStr2 = "";
        RegEx.Pattern = Find;
        RegEx.Global = true;
        _RegExNPos = RegEx.Execute(Src).Item(N).FirstIndex + 1;
        return _RegExNPos;
    }
    public static string RegExNMatch(string Src, string Find, int N = 0)
    {
        string _RegExNMatch = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        dynamic RegM = null;
        string tempStr = "";
        string tempStr2 = "";
        RegEx.Pattern = Find;
        RegEx.Global = true;
        _RegExNMatch = RegEx.Execute(Src).Item(N).Value;
        return _RegExNMatch;
    }
    public static string RegExReplace(string Src, string Find, string Repl)
    {
        string _RegExReplace = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        dynamic RegM = null;
        string tempStr = "";
        string tempStr2 = "";
        RegEx.Pattern = Find;
        RegEx.Global = true;
        _RegExReplace = RegEx.Replace(Src, Repl);
        return _RegExReplace;
    }
    public static dynamic RegExSplit(string szStr, string szPattern)
    {
        dynamic _RegExSplit = null;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        dynamic oAl = null;
        dynamic oRe = null;
        dynamic oMatches = null;
        oRe = RegEx;
        oRe.Pattern = "^(.*)(" + szPattern + ")(.*)$";
        oRe.IgnoreCase = true;
        oRe.Global = true;
        oAl = CreateObject("System.Collections.ArrayList");
        while (true)
        {
            oMatches = oRe.Execute(szStr);
            if (oMatches.Count > 0)
            {
                oAl.Add(oMatches(0).SubMatches(2));
                szStr = oMatches(0).SubMatches(0);
            }
            else
            {
                oAl.Add(szStr);
                break;
            }
        }
        oAl.Reverse();
        _RegExSplit = oAl.ToArray;
        return _RegExSplit;
    }
    public static int RegExSplitCount(string szStr, string szPattern)
    {
        int _RegExSplitCount = 0;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        List<dynamic> T = new List<dynamic>();
        T = RegExSplit(szStr, szPattern);
        _RegExSplitCount = T.Count - 0 + 1;
        return _RegExSplitCount;
    }

}
