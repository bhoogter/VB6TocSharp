using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modConfig;
using static modTextFiles;
using static modUtils;
using static VBExtension;


static class modProjectFiles
{
    // Option Explicit


    public static string VBPModules(string ProjectFile = "")
    {
        string VBPModules = "";
        string S = "";
        dynamic L = null;

        string T = "";

        const string C = "Module=";
        if (ProjectFile == "")
        {
            ProjectFile = vbpFile;
        }
        S = ReadEntireFile(ProjectFile);
        foreach (var iterL in Split(S, vbCrLf))
        {
            L = iterL;
            if (Left(L, Len(C)) == C)
            {
                T = Mid(L, Len(C) + 1);
                if (IsInStr(T, ";"))
                {
                    T = SplitWord(T, 2, ";");
                }
                //If IsInStr(LCase(T), "subclass") Then Stop
                if (LCase(T) == "modlistsubclass.bas")
                {
                    goto NextItem;
                }
                VBPModules = VBPModules + IIf(VBPModules == "", "", vbCrLf) + T;
            }
        NextItem:;
        }
        return VBPModules;
    }

    public static string VBPForms(string ProjectFile = "")
    {
        string VBPForms = "";
        const bool WithExt = true;
        string S = "";
        dynamic L = null;

        string T = "";

        const string C = "Form=";
        if (ProjectFile == "")
        {
            ProjectFile = vbpFile;
        }
        S = ReadEntireFile(ProjectFile);
        foreach (var iterL in Split(S, vbCrLf))
        {
            L = iterL;
            if (Left(L, Len(C)) == C)
            {
                T = Mid(L, Len(C) + 1);
                if (IsInStr(T, ";"))
                {
                    T = SplitWord(T, 1, ";");
                }
                if (!WithExt && Right(T, 4) == ".frm")
                {
                    T = Left(T, Len(T) - 4);
                }
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
                VBPForms = VBPForms + IIf(VBPForms == "", "", vbCrLf) + T;
            }
        NextItem:;
        }
        return VBPForms;
    }

    public static string VBPClasses(string ProjectFile = "", bool ClassNames = false)
    {
        string VBPClasses = "";
        string S = "";
        dynamic L = null;

        string T = "";

        const string C = "Class=";
        if (ProjectFile == "")
        {
            ProjectFile = vbpFile;
        }
        S = ReadEntireFile(ProjectFile);
        foreach (var iterL in Split(S, vbCrLf))
        {
            L = iterL;
            if (Left(L, Len(C)) == C)
            {
                T = Mid(L, Len(C) + 1);
                if (IsInStr(T, ";"))
                {
                    T = SplitWord(T, 2, ";");
                }
                VBPClasses = VBPClasses + IIf(VBPClasses == "", "", vbCrLf) + T;
            }
        NextItem:;
        }
        if (ClassNames)
        {
            VBPClasses = Replace(VBPClasses, ".cls", "");
        }
        return VBPClasses;
    }

    public static string VBPUserControls(string ProjectFile = "")
    {
        string VBPUserControls = "";
        string S = "";
        dynamic L = null;

        string T = "";

        const string C = "UserControl=";
        if (ProjectFile == "")
        {
            ProjectFile = vbpFile;
        }
        S = ReadEntireFile(ProjectFile);
        foreach (var iterL in Split(S, vbCrLf))
        {
            L = iterL;
            if (Left(L, Len(C)) == C)
            {
                T = Mid(L, Len(C) + 1);
                if (IsInStr(T, ";"))
                {
                    T = SplitWord(T, 2, ";");
                }
                VBPUserControls = VBPUserControls + IIf(VBPUserControls == "", "", vbCrLf) + T;
            }
        NextItem:;
        }
        return VBPUserControls;
    }
}
