using Microsoft.VisualBasic;
using System.Collections.Generic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Conversion;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modConfig;
using static modConvertUtils;
using static modUtils;
using static modVB6ToCS;
using static VBExtension;



static class modConvertForm
{
    public static string EventStubs = "";
    public static string FormControlArrays = "";
    public static string Frm2Xml(string F)
    {
        string _Frm2Xml = "";
        List<string> Sp = new List<string>();
        dynamic L = null;
        int I = 0;
        string R = "";
        Sp = new List<string>(Split(F, vbCrLf));
        foreach (var iterL in Sp)
        {
            L = iterL;
            L = Trim(L);
            if (L == "") goto NextLine;
            if (Left(L, 10) == "Attribute " || Left(L, 8) == "VERSION ")
            {
            }
            else if (Left(L, 6) == "Begin ")
            {
                R = R + sSpace(I * SpIndent) + "<item type=\"" + SplitWord(L, 2) + "\" name=\"" + SplitWord(L, 3) + "\">" + vbCrLf;
                I = I + 1;
            }
            else if (L == "End")
            {
                I = I - 1;
                R = R + sSpace(I * SpIndent) + "</item>" + vbCrLf;
            }
            else
            {
                R = R + sSpace(I * SpIndent) + "<prop name=\"" + SplitWord(L, 1, "=") + "\" value=\"" + SplitWord(L, 2, "=", true, true) + "\" />" + vbCrLf;
            }
        NextLine:;
        }
        _Frm2Xml = R;
        return _Frm2Xml;
    }
    public static string FormControls(string Src, string F, bool asLocal = true)
    {
        string _FormControls = "";
        List<string> Sp = new List<string>();
        dynamic L = null;
        int I = 0;
        string R = "";
        string T = "";
        string Nm = "";
        string Ty = "";
        Sp = new List<string>(Split(F, vbCrLf));
        foreach (var iterL in Sp)
        {
            L = iterL;
            L = Trim(L);
            if (L == "") goto NextLine;
            if (Left(L, 6) == "Begin ")
            {
                Ty = SplitWord(L, 2);
                Nm = SplitWord(L, 3);
                switch (Ty)
                {
                    case "VB.Form":
                        break;
                    default:
                        T = Src + ":" + IIf(asLocal, "", Src + ".") + Nm + ":Control:" + Ty;
                        if (Right(R, Len(T)) != T) R = R + vbCrLf + T;
                        break;
                }
            }
        NextLine:;
        }
        _FormControls = R;
        return _FormControls;
    }
    public static string ConvertFormUi(string F, string CodeSection)
    {
        string _ConvertFormUi = "";
        List<string> Stck = new List<string>(); //  TODO: (NOT SUPPORTED) Array ranges not supported: Stck(0 To 100)
        List<string> Sp = new List<string>();
        dynamic L = null;
        int J = 0;
        int K = 0;
        int I = 0;
        string Tag = "";
        string M = "";
        string R = "";
        string Prefix = "";
        Collection Props = null;
        string pK = "";
        string pV = "";
        Sp = new List<string>(Split(F, vbCrLf));
        EventStubs = "";
        FormControlArrays = "";
        for (K = 0; K <= Sp.Count; K += 1)
        {
            L = Trim(Sp[K]);
            if (L == "") goto NextLine;
            if (Left(L, 10) == "Attribute " || Left(L, 8) == "VERSION ")
            {
            }
            else if (Left(L, 6) == "Begin ")
            {
                Props = new Collection();
                J = 0;
                do
                {
                    J = J + 1;
                    M = Trim(Sp[K + J]);
                    if (LMatch(M, "Begin ") || M == "End") break;
                    if (LMatch(M, "BeginProperty "))
                    {
                        Prefix = LCase(Prefix + SplitWord(M, 2) + ".");
                    }
                    else if (LMatch(M, "EndProperty"))
                    {
                        Prefix = Left(Prefix, Len(Prefix) - 1);
                        if (!IsInStr(Prefix, "."))
                        {
                            Prefix = "";
                        }
                        else
                        {
                            Prefix = Left(Prefix, InStrRev(Left(Prefix, Len(Prefix) - 1), "."));
                        }
                    }
                    else
                    {
                        pK = Prefix + LCase(SplitWord(M, 1, "="));
                        pV = ConvertProperty(SplitWord(M, 2, "=", true, true));
                        // TODO: (NOT SUPPORTED): On Error Resume Next
                        Props.Add(pV, pK);
                        // TODO: (NOT SUPPORTED): On Error GoTo 0
                    }
                } while (true);
                K = K + J - 1;
                R = R + sSpace(I * SpIndent) + StartControl(L, Props, LMatch(M, "End"), CodeSection, ref Tag) + vbCrLf;
                I = I + 1;
                Stck[I] = Tag;
            }
            else if (L == "End")
            {
                Props = null;
                Tag = Stck[I];
                I = I - 1;
                if (Tag != "")
                {
                    R = R + sSpace(I * SpIndent) + EndControl(Tag) + vbCrLf;
                }
            }
        NextLine:;
        }
        _ConvertFormUi = R;
        return _ConvertFormUi;
    }
    private static string ConvertProperty(string S)
    {
        string _ConvertProperty = "";
        S = deQuote(S);
        S = DeComment(S);
        _ConvertProperty = S;
        return _ConvertProperty;
    }
    private static string StartControl(string L, Collection Props, bool DoEmpty, string Code, ref string TagType)
    {
        string _StartControl = "";
        string cType = "";
        string oName = "";
        string cName = "";
        string cIndex = "";
        string tType = "";
        bool tCont = false;
        string tDef = "";
        string Features = "";
        string S = "";
        string N = "";
        string M = "";
        string V = "";
        N = vbCrLf;
        TagType = "";
        cType = SplitWord(L, 2);
        oName = SplitWord(L, 3);
        cIndex = cValP(Props, "Index");
        ControlData(cType, ref tType, ref tCont, ref tDef, ref Features);
        if (cIndex != "")
        {
            if (InStr(FormControlArrays, "[" + oName + ",") == 0) FormControlArrays = FormControlArrays + "[" + oName + "," + tType + "]";
            cName = oName + "_" + cIndex;
        }
        else
        {
            cName = oName;
        }
        S = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (tType == "Line" || tType == "Shape" || tType == "Timer")
        {
            return _StartControl;
        }
        else if (tType == "Window")
        {
            S = S + M + "<Window x:Class=\"" + AssemblyName() + ".Forms." + cName + "\"";
            S = S + N + "    xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"";
            S = S + N + "    xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\"";
            S = S + N + "    xmlns:d=\"http://schemas.microsoft.com/expression/blend/2008\"";
            S = S + N + "    xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"";
            S = S + N + "    xmlns:local=\"clr-namespace:" + AssemblyName() + ".Forms\"";
            S = S + N + "    xmlns:usercontrols=\"clr-namespace:" + AssemblyName() + ".UserControls\"";
            S = S + N + "    mc:Ignorable=\"d\"";
            S = S + N + "    Title=" + Quote(cValP(Props, "caption"));
            S = S + M + "    Height=" + Quote(Px(cValP(Props, "clientheight", "0") + 435));
            S = S + M + "    Width=" + Quote(Px(cValP(Props, "clientwidth", "0") + 435));
            S = S + CheckControlEvents("Window", "Form", Code);
            S = S + M + ">";
            S = S + N + " <Grid";
        }
        else if (tType == "GroupBox")
        {
            S = S + "<" + tType;
            S = S + " x:Name=\"" + cName + "\"";
            S = S + " Margin=" + Quote(Px(cValP(Props, "left")) + "," + Px(cValP(Props, "top")) + ",0,0");
            S = S + " Width=" + Quote(Px(cValP(Props, "width")));
            S = S + " Height=" + Quote(Px(cValP(Props, "height")));
            S = S + " VerticalAlignment=\"Top\"";
            S = S + " HorizontalAlignment=\"Left\"";
            S = S + " FontFamily=" + Quote(cValP(Props, "font.name", "Calibri"));
            S = S + " FontSize=" + Quote(cValP(Props, "font.size", "10"));
            S = S + " Header=\"" + cValP(Props, "caption") + "\"";
            S = S + "> <Grid Margin=\"0,-15,0,0\"";
        }
        else if (tType == "Canvas")
        {
            S = S + "<" + tType;
            S = S + " x:Name=\"" + cName + "\"";
            S = S + " Margin=" + Quote(Px(cValP(Props, "left")) + "," + Px(cValP(Props, "top")) + ",0,0");
            S = S + " Width=" + Quote(Px(cValP(Props, "width")));
            S = S + " Height=" + Quote(Px(cValP(Props, "height")));
        }
        else if (tType == "Image")
        {
            S = S + "<" + tType;
            S = S + " x:Name=\"" + cName + "\"";
            S = S + " Margin=" + Quote(Px(cValP(Props, "left")) + "," + Px(cValP(Props, "top")) + ",0,0");
            S = S + " Width=" + Quote(Px(cValP(Props, "width")));
            S = S + " Height=" + Quote(Px(cValP(Props, "height")));
            S = S + " VerticalAlignment=" + Quote("Top");
            S = S + " HorizontalAlignment=" + Quote("Left");
        }
        else
        {
            S = "";
            S = S + "<" + tType;
            S = S + " x:Name=\"" + cName + "\"";
            S = S + " Margin=" + Quote(Px(cValP(Props, "left")) + "," + Px(cValP(Props, "top")) + ",0,0");
            S = S + " Padding=" + Quote("2,2,2,2");
            S = S + " Width=" + Quote(Px(cValP(Props, "width")));
            S = S + " Height=" + Quote(Px(cValP(Props, "height")));
            S = S + " VerticalAlignment=" + Quote("Top");
            S = S + " HorizontalAlignment=" + Quote("Left");
        }
        if (IsInStr(Features, "Font"))
        {
            S = S + " FontFamily=" + Quote(cValP(Props, "font.name", "Calibri"));
            S = S + " FontSize=" + Quote(cValP(Props, "font.size", "10"));
            if (Val(cValP(Props, "font.weight", "400")) > 400) S = S + " FontWeight=" + Quote("Bold");
        }
        if (IsInStr(Features, "Content"))
        {
            S = S + " Content=" + QuoteXML(cValP(Props, "caption") + cValP(Props, "text"));
        }
        if (IsInStr(Features, "Header"))
        {
            S = S + " Content=" + QuoteXML(cValP(Props, "caption") + cValP(Props, "text"));
        }
        V = cValP(Props, "caption") + cValP(Props, "text");
        if (IsInStr(Features, "Text") && V != "")
        {
            S = S + " Text=" + QuoteXML(V);
        }
        V = cValP(Props, "ToolTipText");
        if (IsInStr(Features, "ToolTip") && V != "")
        {
            S = S + " ToolTip=" + Quote(V);
        }
        S = S + CheckControlEvents(tType, cName, Code);
        if (DoEmpty)
        {
            S = S + " />";
            TagType = "";
        }
        else
        {
            S = S + ">";
            TagType = tType;
        }
        _StartControl = S;
        return _StartControl;
    }
    public static string CheckControlEvents(string ControlType, string ControlName, string CodeSection = "")
    {
        string _CheckControlEvents = "";
        string Res = "";
        bool HasClick = false;
        bool HasFocus = false;
        bool HasChange = false;
        bool IsWindow = false;
        HasClick = true;
        HasFocus = !IsInStr("GroupBox", ControlType);
        HasChange = IsInStr("TextBox,ListBox", ControlType);
        IsWindow = ControlType == "Window";
        Res = "";
        Res = Res + CheckEvent("MouseMove", ControlName, ControlType, CodeSection);
        if (HasFocus)
        {
            Res = Res + CheckEvent("GotFocus", ControlName, ControlType, CodeSection);
            Res = Res + CheckEvent("LostFocus", ControlName, ControlType, CodeSection);
            Res = Res + CheckEvent("KeyDown", ControlName, ControlType, CodeSection);
            Res = Res + CheckEvent("KeyUp", ControlName, ControlType, CodeSection);
        }
        if (HasClick)
        {
            Res = Res + CheckEvent("Click", ControlName, ControlType, CodeSection);
            Res = Res + CheckEvent("DblClick", ControlName, ControlType, CodeSection);
        }
        if (HasChange)
        {
            Res = Res + CheckEvent("Change", ControlName, ControlType, CodeSection);
        }
        if (IsWindow)
        {
            Res = Res + CheckEvent("Load", ControlName, ControlType, CodeSection);
            Res = Res + CheckEvent("Unload", ControlName, ControlType, CodeSection);
            // Res = Res & CheckEvent(__S1, ControlName, ControlType, CodeSection)
        }
        _CheckControlEvents = Res;
        return _CheckControlEvents;
    }
    public static string CheckEvent(string EventName, string ControlName, string ControlType, string CodeSection = "")
    {
        string _CheckEvent = "";
        string Search = "";
        string Target = "";
        string N = "";
        int L = 0;
        string V = "";
        N = ControlName + "_" + EventName;
        Search = " " + N + "(";
        Target = EventName;
        switch (EventName)
        {
            case "DblClick":
                Target = "MouseDoubleClick";
                break;
            case "Change":
                if (ControlType == "TextBox") Target = "TextChanged";
                break;
            case "Load":
                Target = "Loaded";
                break;
            case "Unload":
                Target = "Unloaded";
                break;
        }
        L = InStr(1, CodeSection, Search, vbTextCompare);
        if (L > 0)
        {
            V = Mid(CodeSection, L + 1, Len(N)); // Get exact capitalization from source....
            _CheckEvent = " " + Target + "=\"" + V + "\"";
        }
        else
        {
            _CheckEvent = "";
        }
        return _CheckEvent;
    }
    public static string EndControl(string tType)
    {
        string _EndControl = "";
        switch (tType)
        {
            case "Line":
            case "Shape":
            case "Timer":
                _EndControl = "";
                break;
            case "Window":
                _EndControl = " </Grid>" + vbCrLf + "</Window>";
                break;
            case "GroupBox":
                _EndControl = "</Grid> </GroupBox>";
                break;
            default:
                _EndControl = "</" + tType + ">";
                break;
        }
        return _EndControl;
    }
    public static bool IsEvent(string Str)
    {
        bool _IsEvent = false;
        _IsEvent = EventStub(Str) != "";
        return _IsEvent;
    }
    public static string EventStub(string fName)
    {
        string _EventStub = "";
        string S = "";
        string C = "";
        string K = "";
        C = SplitWord(fName, 1, "_");
        K = SplitWord(fName, 2, "_");
        switch (K)
        {
            case "Click":
            case "DblClick":
            case "Load":
            case "GotFocus":
            case "LostFocus":
                S = "private void " + fName + "(object sender, RoutedEventArgs e) { " + fName + "(); }" + vbCrLf;
                break;
            case "Change":
                S = "private void " + C + "_Change(object sender, System.Windows.Controls.TextChangedEventArgs e) { " + fName + "(); }" + vbCrLf;
                break;
            case "QueryUnload":
                S = "private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) { int c = 0, u = 0 ;  " + fName + "(out c, ref u); e.Cancel = c != 0;  }" + vbCrLf;
                // V = __S1 & FName & __S2
                break;
            case "Validate":
            case "Unload":
                // V = __S1 & FName & __S2
                break;
            case "KeyDown":
            case "KeyUp":
            case "KeyPress":
                break;
            case "MouseMove":
            case "MouseDown":
            case "MouseUp":
                break;
        }
        _EventStub = S;
        return _EventStub;
    }

}
