using System.Collections.Generic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modConfig;
using static modTextFiles;
using static modUtils;
using static VBExtension;



static class modProjectFiles
{
    public static string VBPCode(string ProjectFile = "")
    {
        string _VBPCode = "";
        _VBPCode = VBPModules() + vbCrLf + VBPForms() + vbCrLf + VBPClasses() + vbCrLf + VBPUserControls();
        return _VBPCode;
    }
    public static string VBPModules(string ProjectFile = "")
    {
        string _VBPModules = "";
        string S = "";
        dynamic L = null;
        string T = "";
        string C = "Module=";
        if (ProjectFile == "") ProjectFile = vbpFile;
        S = ReadEntireFile(ProjectFile);
        foreach (var iterL in new List<string>(Split(S, vbCrLf)))
        {
            L = iterL;
            if (Left(L, Len(C)) == C)
            {
                T = Mid(L, Len(C) + 1);
                if (IsInStr(T, ";")) T = SplitWord(T, 2, ";");
                // If IsInStr(LCase(T), __S1) Then Stop
                if (LCase(T) == "modlistsubclass.bas") goto NextItem;
                _VBPModules = _VBPModules + IIf(_VBPModules == "", "", vbCrLf) + T;
            }
        NextItem:;
        }
        return _VBPModules;
    }
    public static string VBPForms(string ProjectFile = "")
    {
        string _VBPForms = "";
        bool WithExt = true;
        string S = "";
        dynamic L = null;
        string T = "";
        string C = "Form=";
        if (ProjectFile == "") ProjectFile = vbpFile;
        S = ReadEntireFile(ProjectFile);
        foreach (var iterL in new List<string>(Split(S, vbCrLf)))
        {
            L = iterL;
            if (Left(L, Len(C)) == C)
            {
                T = Mid(L, Len(C) + 1);
                if (IsInStr(T, ";")) T = SplitWord(T, 1, ";");
                if (!WithExt && Right(T, 4) == ".frm") T = Left(T, Len(T) - 4);
                switch (LCase(T))
                {
                    case "faxtest":
                        T = "FaxPO";
                        break;
                    case "frmpos":
                        T = "frmCashRegister";
                        break;
                    case "frmposquantity":
                        T = "frmCashRegisterQuantity";
                        break;
                    case "calendarinst":
                        T = "CalendarInstr";
                        break;
                    case "frmedi":
                        T = "frmAshleyEDIItemAlign";
                        break;
                    case "frmpracticefiles":
                        T = "PracticeFiles";
                        break;
                    case "txttextselect":
                        T = "frmSelectText";
                        break;
                }
                _VBPForms = _VBPForms + IIf(_VBPForms == "", "", vbCrLf) + T;
            }
        NextItem:;
        }
        return _VBPForms;
    }
    public static string VBPClasses(string ProjectFile = "", bool ClassNames = false)
    {
        string _VBPClasses = "";
        string S = "";
        dynamic L = null;
        string T = "";
        string C = "Class=";
        if (ProjectFile == "") ProjectFile = vbpFile;
        S = ReadEntireFile(ProjectFile);
        foreach (var iterL in new List<string>(Split(S, vbCrLf)))
        {
            L = iterL;
            if (Left(L, Len(C)) == C)
            {
                T = Mid(L, Len(C) + 1);
                if (IsInStr(T, ";")) T = SplitWord(T, 2, ";");
                _VBPClasses = _VBPClasses + IIf(_VBPClasses == "", "", vbCrLf) + T;
            }
        NextItem:;
        }
        if (ClassNames) _VBPClasses = Replace(_VBPClasses, ".cls", "");
        return _VBPClasses;
    }
    public static string VBPUserControls(string ProjectFile = "")
    {
        string _VBPUserControls = "";
        string S = "";
        dynamic L = null;
        string T = "";
        string C = "UserControl=";
        if (ProjectFile == "") ProjectFile = vbpFile;
        S = ReadEntireFile(ProjectFile);
        foreach (var iterL in new List<string>(Split(S, vbCrLf)))
        {
            L = iterL;
            if (Left(L, Len(C)) == C)
            {
                T = Mid(L, Len(C) + 1);
                if (IsInStr(T, ";")) T = SplitWord(T, 2, ";");
                _VBPUserControls = _VBPUserControls + IIf(_VBPUserControls == "", "", vbCrLf) + T;
            }
        NextItem:;
        }
        return _VBPUserControls;
    }

}
