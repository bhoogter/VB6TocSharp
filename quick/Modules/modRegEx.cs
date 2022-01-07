using System;
using System.Collections.Generic;
using static Microsoft.VisualBasic.Interaction;
using static VBExtension;



static class modRegEx
{
    public static dynamic mRegEx = null;
    private static dynamic RegEx
    {
        get
        {
            dynamic _RegEx;
            if (mRegEx == null) { mRegEx = CreateObject("vbscript.regexp"); mRegEx.Global = true; }
            _RegEx = mRegEx;
            return _RegEx;
        }
    }

    public static bool RegExTest(string Src, string Find)
    {
        bool _RegExTest = false;
        try
        {
            RegEx.Pattern = Find;
            _RegExTest = RegEx.test(Src);
        }
        catch (Exception e) { }
        return _RegExTest;
    }
    public static int RegExCount(string Src, string Find)
    {
        int _RegExCount = 0;
        try
        {
            RegEx.Pattern = Find;
            RegEx.Global = true;
            _RegExCount = RegEx.Execute(Src).Count;
        }
        catch (Exception e) { }
        return _RegExCount;
    }
    public static int RegExNPos(string Src, string Find, int N = 0)
    {
        int _RegExNPos = 0;
        try
        {
            dynamic RegM = null;
            string tempStr = "";
            string tempStr2 = "";
            RegEx.Pattern = Find;
            RegEx.Global = true;
            _RegExNPos = RegEx.Execute(Src).Item(N).FirstIndex + 1;
        }
        catch (Exception e) { }
        return _RegExNPos;
    }
    public static string RegExNMatch(string Src, string Find, int N = 0)
    {
        string _RegExNMatch = "";
        try
        {
            dynamic RegM = null;
            string tempStr = "";
            string tempStr2 = "";
            RegEx.Pattern = Find;
            RegEx.Global = true;
            _RegExNMatch = RegEx.Execute(Src).Item(N).Value;
        }
        catch (Exception e) { }
        return _RegExNMatch;
    }
    public static string RegExReplace(string Src, string Find, string Repl)
    {
        string _RegExReplace = "";
        try
        {
            dynamic RegM = null;
            string tempStr = "";
            string tempStr2 = "";
            RegEx.Pattern = Find;
            RegEx.Global = true;
            _RegExReplace = RegEx.Replace(Src, Repl);
        }
        catch (Exception e) { }
        return _RegExReplace;
    }
    public static List<string> RegExSplit(string szStr, string szPattern)
    {
        List<string> _RegExSplit = null;
        try
        {
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
        }
        catch (Exception e) { }
        return _RegExSplit;
    }
    public static int RegExSplitCount(string szStr, string szPattern)
    {
        int _RegExSplitCount = 0;
        try
        {
            List<string> T;
            T = RegExSplit(szStr, szPattern);
            _RegExSplitCount = T.Count - 0 + 1;
        }
        catch (Exception e) { }
        return _RegExSplitCount;
    }
}
