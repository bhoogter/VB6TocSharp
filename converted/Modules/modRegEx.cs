using System.Collections.Generic;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static VBExtension;


static class modRegEx
{
    // Option Explicit
    private static dynamic mRegEx = null;


    static dynamic RegEx
    {
        get
        {
            dynamic RegEx;
            if (mRegEx == null)
            {
                mRegEx = CreateObject("vbscript.regexp");
                mRegEx.Global = true;
            }
            RegEx = mRegEx;

            return RegEx;
        }
    }


    public static bool RegExTest(string Src, string Find)
    {
        bool RegExTest = false;
        // TODO (not supported): On Error Resume Next
        RegEx.Pattern = Find;
        RegExTest = RegEx.Test(Src);
        return RegExTest;
    }

    public static int RegExCount(string Src, string Find)
    {
        int RegExCount = 0;
        // TODO (not supported): On Error Resume Next
        RegEx.Pattern = Find;
        RegEx.Global = true;
        RegExCount = RegEx.Execute(Src).Count;
        return RegExCount;
    }

    public static int RegExNPos(string Src, string Find, int N = 0)
    {
        int RegExNPos = 0;
        // TODO (not supported): On Error Resume Next
        dynamic RegM = null;
        string tempStr = "";
        string tempStr2 = "";

        RegEx.Pattern = Find;
        RegEx.Global = true;
        RegExNPos = RegEx.Execute(Src).Item(N).FirstIndex + 1;
        return RegExNPos;
    }

    public static string RegExNMatch(string Src, string Find, int N = 0)
    {
        string RegExNMatch = "";
        // TODO (not supported): On Error Resume Next
        dynamic RegM = null;
        string tempStr = "";
        string tempStr2 = "";

        RegEx.Pattern = Find;
        RegEx.Global = true;
        RegExNMatch = RegEx.Execute(Src).Item(N).Value;
        return RegExNMatch;
    }

    public static string RegExReplace(string Src, string Find, string Repl)
    {
        string RegExReplace = "";
        // TODO (not supported): On Error Resume Next
        dynamic RegM = null;
        string tempStr = "";
        string tempStr2 = "";

        RegEx.Pattern = Find;
        RegEx.Global = true;
        RegExReplace = RegEx.Replace(Src, Repl);
        return RegExReplace;
    }

    public static dynamic RegExSplit(string szStr, string szPattern)
    {
        dynamic RegExSplit = null;
        // TODO (not supported): On Error Resume Next
        dynamic oAl = null;
        dynamic oRe = null;
        dynamic oMatches = null;

        oRe = RegEx;
        oRe.Pattern = "^(.*)(" + szPattern + ")(.*)$";
        oRe.IgnoreCase = true;
        oRe.Global = true;
        oAl = CreateObject("System.Collections.ArrayList");

        do
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
        RegExSplit = oAl.ToArray;
        return RegExSplit;
    }

    public static int RegExSplitCount(string szStr, string szPattern)
    {
        int RegExSplitCount = 0;
        // TODO (not supported): On Error Resume Next
        List<dynamic> T = new List<dynamic> { }; // TODO - Specified Minimum Array Boundary Not Supported:   Dim T()

        T = RegExSplit(szStr, szPattern);
        RegExSplitCount = UBound(T) - LBound(T) + 1;
        return RegExSplitCount;
    }
}
