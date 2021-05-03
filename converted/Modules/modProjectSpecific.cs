using static Microsoft.VisualBasic.Strings;
using static modRegEx;
using static modUtils;


static class modProjectSpecific
{
    // Option Explicit


    public static string ProjectSpecificPostCodeLineConvert(string Str)
    {
        string ProjectSpecificPostCodeLineConvert = "";
        string S = "";

        S = Str;

        //  If IsInStr(S, "!C == null") Then Stop

        // Some patterns we dont use or didn't catch in lint...
        if (IsInStr(S, "DisposeDA"))
        {
            S = Replace(S, "DisposeDA", "// DisposeDA");
        }
        if (IsInStr(S, "MousePointer = vbNormal"))
        {
            S = Replace(S, "MousePointer = vbNormal", "MousePointer = vbDefault");
        }

        // We use decimal, not double
        if (IsInStr(S, "Val("))
        {
            S = Replace(S, "Val( ", "ValD(");
        }

        // Bad pattern combination
        if (RegExTest(S, "\\(!" + patToken + " == null\\)"))
        {
            S = Replace(S, "!", "", 1);
            S = Replace(S, "==", "!=", 1);
        }

        // False ref entries...
        if (IsInStr(S, "IsIn("))
        {
            S = Replace(S, "ref ", "");
        }
        if (IsInStr(S, "POMode("))
        {
            S = Replace(S, "ref ", "");
        }
        if (IsInStr(S, "OrderMode("))
        {
            S = Replace(S, "ref ", "");
        }
        if (IsInStr(S, "InvenMode("))
        {
            S = Replace(S, "ref ", "");
        }
        if (IsInStr(S, "ReportsMode("))
        {
            S = Replace(S, "ref ", "");
        }
        if (IsInStr(S, "SetButtonImage("))
        {
            S = Replace(S, "ref ", "");
            S = Replace(S, ".DefaultProperty", "");
        }
        if (IsInStr(S, "EnableFrame"))
        {
            S = Replace(S, "ref ", "");
        }
        S = Replace(S, " && BackupType.", " & BackupType.");

        // Common Mistake Functions...
        if (IsInStr(S, "StoreSettings."))
        {
            S = Replace(S, "StoreSettings.", "StoreSettings().");
        }

        // etc
        if (IsInStr(S, ".hwnd"))
        {
            S = Replace(S, ".hwnd", ".hWnd()");
        }
        if (IsInStr(S, "SetCustomFrame"))
        {
            S = "";
        }
        if (IsInStr(S, "RemoveCustomFrame"))
        {
            S = "";
        }
        S = Replace(S, "VbMsgBoxResult", "MsgBoxResult");

        const string TokenBreak = "[ ,)]";
        S = RegExReplace(S, "InventFolder(" + TokenBreak + ")", "InventFolder()$1");
        S = RegExReplace(S, "PXFolder(" + TokenBreak + ")", "InventFolder()$1");
        S = RegExReplace(S, "FXFolder(" + TokenBreak + ")", "InventFolder()$1");
        S = RegExReplace(S, "InventFolder(" + TokenBreak + ")", "InventFolder()$1");
        S = RegExReplace(S, "IsDevelopment(" + TokenBreak + ")", "IsDevelopment()$1");

        ProjectSpecificPostCodeLineConvert = S;
        return ProjectSpecificPostCodeLineConvert;
    }
}
