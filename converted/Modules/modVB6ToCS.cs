using System;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Conversion;
using static Microsoft.VisualBasic.Strings;
using static modConfig;
using static modProjectFiles;
using static modRegEx;
using static modSubTracking;
using static modUtils;
using static VBExtension;


static class modVB6ToCS
{
    // Option Explicit


    public static string ConvertDefaultDefault(string DType)
    {
        string ConvertDefaultDefault = "";
        switch (DType)
        {
            case "Integer":
                ConvertDefaultDefault = 0;
                break;
            case "Date":
                ConvertDefaultDefault = "DateTime.MinValue";
                break;
            case "String":
                ConvertDefaultDefault = "\"\"";
                break;
            case "Boolean":
                ConvertDefaultDefault = "false";
                break;
            default:
                ConvertDefaultDefault = "null";
                break;
        }
        return ConvertDefaultDefault;
    }

    public static string ConvertDataType(string S)
    {
        string ConvertDataType = "";
        switch (S)
        {
            case "Object":
                ConvertDataType = DefaultDataType;
                break;
            case "Form":
                ConvertDataType = "Window";
                break;
            case "String":
                ConvertDataType = "string";
                break;
            case "String()":
                ConvertDataType = "List<string>";
                break;
            case "Long":
                ConvertDataType = "int";
                break;
            case "Integer":
                ConvertDataType = "int";
                break;
            case "Double":
                ConvertDataType = "decimal";
                break;
            case "Variant":
                ConvertDataType = "object";
                break;
            case "Byte":
                ConvertDataType = "byte";
                break;
            case "Boolean":
                ConvertDataType = "bool";
                break;
            case "Currency":
                ConvertDataType = "decimal";
                break;
            case "VbTriState":
                ConvertDataType = "vbTriState";
                break;
            case "Collection":
                ConvertDataType = "Collection";
                break;
            case "TSPNode":
                ConvertDataType = "TSPNode";
                break;
            case "TSPNetwork":
                ConvertDataType = "TSPNetwork";
                break;
            case "FindResults":
                ConvertDataType = "FindResults";
                break;
            case "Pushpin":
                ConvertDataType = "Pushpin";
                break;
            case "Map":
                ConvertDataType = "Map";

                break;
            case "Node":
                ConvertDataType = "TreeViewItem";

                break;
            case "Recordset":
                ConvertDataType = "Recordset";
                break;
            case "Connection":
                ConvertDataType = "Connection";
                break;
            case "ADODB.Error":
                ConvertDataType = "ADODB.Error";
                break;
            case "ADODB.EventStatusEnum":
                ConvertDataType = "ADODB.EventStatusEnum";
                break;
            case "SpeechLib.SpeechEngineConfidence":
                ConvertDataType = DefaultDataType;

                break;
            case "Date":
                ConvertDataType = "DateTime";
                break;
            case "VbMsgBoxResult":
                ConvertDataType = S;
                break;
            case "XCTransaction2.XChargeTransaction":
                ConvertDataType = S;
                break;
            case "PictureBox":
                ConvertDataType = S;
                break;
            case "MSCommLib.MSComm":
                ConvertDataType = S;

                break;
            default:
                if (IsInStr(VBPClasses(ClassNames: true), S))
                {
                    ConvertDataType = S;
                }
                else
                {
                    ConvertDataType = S;
                    Console.WriteLine("Unknown Data Type: " + S);
                }
                break;
        }
        return ConvertDataType;
    }

    public static dynamic ControlData(string cType, out string Name, out bool Cont, out string Def, out string Features)
    {
        dynamic ControlData = null;
        Cont = false;
        Def = "Caption";
        switch (cType)
        {
            case "VB.Form":
                Name = "Window";
                Cont = true;
                break;
            case "VB.MDIForm":
                Name = "Window";
                Cont = true;
                Cont = true;

                break;
            case "VB.PictureBox":
                Name = "Image";
                Cont = true;
                Def = "Picture";
                Features = "Tooltiptext";
                break;
            case "VB.Label":
                Name = "Label";
                Features = "";
                Features = "Font,Content,Tooltiptext";
                break;
            case "VB.TextBox":
                Name = "TextBox";
                Def = "Text";
                Features = "Font,Text,Tooltiptext";
                break;
            case "VB.Frame":
                Name = "GroupBox";
                Features = "Tooltiptext";
                break;
            case "VB.CommandButton":
                Name = "Button";
                Features = "Font,Content,Tooltiptext";
                break;
            case "VB.CheckBox":
                Name = "CheckBox";
                Features = "Font,Content,Tooltiptext";
                break;
            case "VB.OptionButton":
                Name = "RadioButton";
                Features = "Font,Content,Tooltiptext";
                break;
            case "VB.ComboBox":
                Name = "ComboBox";
                Def = "Text";
                Features = "Font,Text,Tooltiptext";
                break;
            case "VB.ListBox":
                Name = "ListBox";
                Def = "Text";
                Features = "Font,Tooltiptext";
                break;
            case "VB.HScrollBar":
                Name = "ScrollBar";
                Def = "Value";
                Features = "";
                break;
            case "VB.VScrollBar":
                Name = "ScrollBar";
                Def = "Value";
                Features = "";
                break;
            case "VB.Timer":
                Name = "Timer";
                Def = "Enabled";
                Features = "";
                break;
            case "VB.DriveListBox":
                Name = "usercontrols:DriveListBox";
                Def = "Path";
                Features = "";
                break;
            case "VB.DirListBox":
                Name = "usercontrols:DirListBox";
                Def = "Path";
                Features = "";
                break;
            case "VB.FileListBox":
                Name = "usercontrols:FileListBox";
                Def = "Path";
                Features = "";
                break;
            case "VB.Shape":
                Name = "Shape";
                Def = "Visible";
                Features = "";
                break;
            case "VB.Line":
                Name = "Line";
                Def = "Visible";
                Features = "";
                break;
            case "VB.Image":
                Name = "Image";
                Def = "Picture";
                Features = "Tooltiptext";
                break;
            case "VB.Data":
                Name = "Data";
                Def = "DataSource";
                Features = "";
                break;
            case "VB.OLE":
                Name = "OLE";
                Def = "OLE";
                Features = "";

                break;
            case "VB.Menu":
                Name = "Menu";

                // MS Windows Common Controls 6.0
                break;
            case "MSComctlLib.TabStrip":
                break;
            case "MSComctlLib.ToolBar":
                break;
            case "MSComctlLib.StatusBar":
                Name = "StatusBar";
                Def = "Text";
                Features = "Tooltiptext";
                break;
            case "MSComctlLib.ProgressBar":
                Name = "ProgressBar";
                Def = "Value";
                Features = "Tooltiptext";
                break;
            case "MSComctlLib.TreeView":
                Name = "TreeView";
                Features = "Tooltiptext";
                break;
            case "MSComctlLib.ListView":
                Name = "ListView";
                Features = "Tooltiptext";
                break;
            case "MSComctlLib.ImageList":
                Name = "ImageList";
                Features = "Tooltiptext";
                break;
            case "MSComctlLib.Slider":
                Name = "Slider";
                break;
            case "MSComctlLib.ImageCombo":
                // MS Windows Common Controls-2 6.0
                //    Case "MSComCtl2.Animation":
                break;
            case "MSComCtl2.UpDown":
                Name = "usercontrols:UpDown";
                break;
            case "MSComCtl2.DTPicker":
                Name = "DatePicker";
                break;
            case "MSComCtl2.MonthView":
                Name = "DatePicker";
                break;
            case "MSComCtl2.FlatScrollBar":
                Name = "ScrollBar";

                break;
            case "MSComDlg.CommonDialog":
                Name = "Label";
                break;
            case "MSFlexGridLib.MSFlexGrid":
                Name = "usercontrols:FlexGrid";
                break;
            case "MSDBGrid.DBGrid":
                Name = "DataGrid";
                break;
            case "TabDlg.SSTab":
                Name = "TabControl";
                break;
            case "RichTextLib.RichTextBox":
                Name = "TextBlock";
                break;
            case "InetCtlsObjects.Inet":
                Name = "INet";
                break;
            case "MSCommLib.MSComm":
                Name = "MSComm";
                break;
            case "MSWinsockLib.Winsock":
                Name = "Winsock";

                break;
            case "WinCDS.UGridIO":
                Name = "UGridIO";
                break;
            case "WinCDS.CandyButton":
                Name = "Button";
                break;
            case "WinCDS.ucPBar":
                Name = "ProgressBar";
                break;
            case "WinCDS.PrinterSelector":
                Name = "Label";
                break;
            case "WinCDS.RichTextBoxNew":
                Name = "TextBlock";
                break;
            case "WinCDS.MaskedPicture":
                Name = "Image";

                break;
            case "VJCZIPLib.VjcZip":
                Name = "Label";
                break;
            case "MSChart20Lib.MSChart":
                Name = "Label";
                break;
            case "MapPointCtl.MappointControl":
                Name = "Label";

                break;
            case "LaVolpeAlphaImg.AlphaImgCtl":
                Name = "Image";
                break;
            case "GIF89LibCtl.Gif89a":
                Name = "Image";

                break;
            default:
                Console.WriteLine("Unknown Control Type: " + cType);
                Name = "Label";
                break;
        }
        return ControlData;
    }

    public static dynamic ConvertVb6Specific(string S, out bool Complete)
    {
        dynamic ConvertVb6Specific = null;
        string W = "";
        string R = "";


        switch (Trim(S))
        {
            case "Array()":
                S = "new List<dynamic>()";
                break;
            case "App.Path":
                S = "AppDomain.CurrentDomain.BaseDirectory";
                break;
        }

        Complete = false;
        W = RegExNMatch(Trim(S), patToken);
        R = SplitWord(Trim(S), 2, " ", true, true);
        switch (W)
        {
            case "True":
                Complete = true;
                S = "true";
                break;
            case "False":
                Complete = true;
                S = "false";
                break;
            case "Me":
                Complete = true;
                S = "this";
                break;
            case "Nothing":
                Complete = true;
                S = "null";
                break;
            case "vbTrue":
                Complete = true;
                S = "vbTriState.vbTrue";
                break;
            case "vbFalse":
                Complete = true;
                S = "vbTriState.vbFalse";
                break;
            case "vbUseDefault":
                Complete = true;
                S = "vbTriState.vbUseDefault";
                break;
            case "Date":
                Complete = true;
                S = "DateTime.Today;";
                break;
            case "Now":
                Complete = true;
                S = "DateTime.Now;";
                break;
            case "Kill":
                S = "File.Delete(" + R + ");";
                break;
            case "FreeFile":
                S = "FreeFile();";
                break;
            case "Open":
                S = "VBOpenFile(" + Replace(SplitWord(R, 2, " As "), "#", "") + ", " + SplitWord(R, 1, " For ") + ");";
                break;
            case "Print":
                S = "VBWriteFile(" + Replace(SplitWord(R, 1, ","), "#", "") + ", " + Replace(SplitWord(R, 2, ", ", true, true), ";", ",") + ");";
                break;
            case "Close":
                S = "VBCloseFile(" + Replace(R, "#", "") + ");";
                break;
            case "New":
                Complete = true;
                S = "new " + R + "();";
                break;
            case "vbAlignLeft":
                S = "AlignConstants.vbAlignLeft";
                break;
            case "vbAlignRight":
                S = "AlignConstants.vbAlignRight";
                break;
            case "vbAlignTop":
                S = "AlignConstants.vbAlignTop";
                break;
            case "vbAlignBottom":
                S = "AlignConstants.vbAlignBottom";
                break;
            case "RaiseEvent":
                Complete = true;
                W = RegExNMatch(R, patToken);
                R = Mid(R, Len(W) + 1);
                if (R == "")
                {
                    R = "()";
                }
                S = "event" + W + "?.Invoke" + R + ";";
                break;
            case "ReDim":
                Complete = true;
                bool RedimPres = false;
                string RedimVar = "";
                string RedimTyp = "";
                string RedimTmp = "";
                string RedimMax = "";
                string RedimIter = "";

                if (tLMatch(R, "Preserve "))
                {
                    R = Trim(tMid(R, 10));
                    RedimPres = true;
                }

                RedimVar = RegExNMatch(R, patToken);
                RedimTyp = ConvertDataType(SubParam(RedimVar).asType);
                R = Trim(Replace(R, RedimVar, ""));
                if (tLeft(R, 1) == "(")
                {
                    R = Mid(Trim(R), 2);
                }
                RedimMax = Val(nextBy(R, ")"));
                RedimTmp = RedimVar + "_" + Random() + "_tmp";
                RedimIter = "redim_iter_" + Random();
                S = "";
                S = S + "List<" + RedimTyp + "> " + RedimTmp + " = new List<" + RedimTyp + ">();" + vbCrLf;

                S = S + "for (int " + RedimIter + "=0;i<" + RedimMax + ";" + RedimIter + "++) {";
                if (RedimPres)
                {
                    S = S + RedimVar + ".Add(" + RedimIter + "<" + RedimVar + ".Count ? " + RedimVar + "(" + RedimIter + ") : " + ConvertDefaultDefault(SubParam(RedimVar).asType) + ");";
                }
                else
                {
                    S = S + RedimVar + ".Add(" + ConvertDefaultDefault(SubParam(RedimVar).asType) + ");";
                }
                S = S + "}";
                break;
        }

        if (IsInStr(S, ".Print "))
        {
            if (Right(S, 1) == ";")
            {
                S = Replace(S, ".Print ", ".PrintNNL ");
                S = Left(S, Len(S) - 1);
            }
            S = Replace(S, ";", ",");
        }

        ConvertVb6Specific = S;
        return ConvertVb6Specific;
    }

    public static string ConvertVb6Syntax(string S)
    {
        string ConvertVb6Syntax = "";
        string W = "";
        string R = "";

        W = RegExNMatch(Trim(S), patToken);
        R = SplitWord(Trim(S), 2, " ", true, true);
        switch (W)
        {
            case "Open":
                S = "VBOpenFile(" + Replace(SplitWord(R, 2, " As "), "#", "") + ", " + SplitWord(R, 1, " For ") + ")";
                break;
            case "Print":
                S = "VBWriteFile(" + Replace(SplitWord(R, 1, ","), "#", "") + ", " + Replace(SplitWord(R, 2, ", ", true, true), ";", ",") + ")";
                break;
            case "Input":
                S = "VBReadFile(" + Replace(SplitWord(R, 1, ","), "#", "") + ", " + Replace(SplitWord(R, 2, ", ", true, true), ";", ",") + ")";
                break;
            case "Line":
                S = "VBReadFileLine(" + Replace(SplitWord(R, 1, ","), "#", "") + ", " + Replace(SplitWord(R, 2, ", ", true, true), ";", ",") + ")";
                break;
            case "Close":
                S = "VBCloseFile(" + Replace(R, "#", "") + ")";
                break;
            case "New":
                S = "new " + R + "()";
                break;
            case "RaiseEvent":
                W = RegExNMatch(R, patToken);
                R = Mid(R, Len(W) + 1);
                if (R == "")
                {
                    R = "()";
                }
                S = "event" + W + "?.Invoke" + R;
                break;
        }

        ConvertVb6Syntax = S;
        return ConvertVb6Syntax;
    }
}
