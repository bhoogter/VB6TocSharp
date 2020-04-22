using VB6 = Microsoft.VisualBasic.Compatibility.VB6;
using System.Runtime.InteropServices;
using static VBExtension;
using static VBConstants;
using Microsoft.VisualBasic;
using System;
using System.Windows;
using System.Windows.Controls;
using static System.DateTime;
using static System.Math;
using static Microsoft.VisualBasic.Globals;
using static Microsoft.VisualBasic.Collection;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Conversion;
using static Microsoft.VisualBasic.DateAndTime;
using static Microsoft.VisualBasic.ErrObject;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Financial;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static Microsoft.VisualBasic.VBMath;
using System.Collections.Generic;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ColorConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.DrawStyleConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.FillStyleConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.GlobalModule;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.Printer;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterCollection;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterObjectConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ScaleModeConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.SystemColorConstants;
using ADODB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using VB2CS.Forms;
using static modUtils;
using static modConvert;
using static modProjectFiles;
using static modTextFiles;
using static modRegEx;
using static frmTest;
using static modConvertForm;
using static modSubTracking;
using static modVB6ToCS;
using static modUsingEverything;
using static modSupportFiles;
using static modConfig;
using static modRefScan;
using static modConvertUtils;
using static modControlProperties;
using static modProjectSpecific;
using static modINI;
using static modLinter;
using static VB2CS.Forms.frm;
using static VB2CS.Forms.frmConfig;


static class modConvertForm {
// Option Explicit
private static string EventStubs = "";


public static string Frm2Xml(string F) {
  string Frm2Xml = "";
  dynamic Sp = null;
  dynamic L = null;
  int I = 0;

  string R = "";

  Sp = Split(F, vbCrLf);

  foreach(var L in Sp) {
    L = Trim(L);
    if (L == "") {
goto NextLine;
    }
    if (Left(L, 10) == "Attribute " || Left(L, 8) == "VERSION ") {
    } else if (Left(L, 6) == "Begin ") {
      R = R + sSpace(I * SpIndent) + "<item type=\"" + SplitWord(L, 2) + "\" name=\"" + SplitWord(L, 3) + "\">" + vbCrLf;
      I = I + 1;
    } else if (L == "End") {
      I = I - 1;
      R = R + sSpace(I * SpIndent) + "</item>" + vbCrLf;
    } else {
      R = R + sSpace(I * SpIndent) + "<prop name=\"" + SplitWord(L, 1, "=") + "\" value=\"" + SplitWord(L, 2, "=", true, true) + "\" />" + vbCrLf;
    }
NextLine:
  }
  Frm2Xml = R;
  return Frm2Xml;
}

public static dynamic FormControls(string Src, string F, bool asLocal= true) {
  dynamic FormControls = null;
  dynamic Sp = null;
  dynamic L = null;
  int I = 0;

  string R = "";
  string T = "";

  string Nm = "";
  string Ty = "";

  Sp = Split(F, vbCrLf);

  foreach(var L in Sp) {
    L = Trim(L);
    if (L == "") {
goto NextLine;
    }
    if (Left(L, 6) == "Begin ") {
      Ty = SplitWord(L, 2);
      Nm = SplitWord(L, 3);
      switch(Ty) {
        case "VB.Form":
          break;
        default:
          T = Src + ":" + IIf(asLocal, "", Src + ".") + Nm + ":Control:" + Ty;
          if (Right(R, Len(T)) != T) {
            R = R + vbCrLf + T;
          }
break;
}
    }
NextLine:
  }
  FormControls = R;
  return FormControls;
}

public static string ConvertFormUi(string F, string CodeSection) {
  string ConvertFormUi = "";
  List<dynamic> Stck = new List<dynamic> (new dynamic[101]);

  dynamic Sp = null;
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

  Sp = Split(F, vbCrLf);

  EventStubs = "";

  for(K=LBound(Sp); K<UBound(Sp); K++) {
    L = Trim(Sp(K));
    if (L == "") {
goto NextLine;
    }

    if (Left(L, 10) == "Attribute " || Left(L, 8) == "VERSION ") {
    } else if (Left(L, 6) == "Begin ") {
      Props = new Collection();;
      J = 0;
      do {
        J = J + 1;
        M = Trim(Sp(K + J));
        if (LMatch(M, "Begin ") || M == "End") {
          break;
        }

        if (LMatch(M, "BeginProperty ")) {
          Prefix = LCase(Prefix + SplitWord(M, 2) + ".");
        } else if (LMatch(M, "EndProperty")) {
          Prefix = Left(Prefix, Len(Prefix) - 1);
          if (!IsInStr(Prefix, ".")) {
            Prefix = "";
          } else {
            Prefix = Left(Prefix, InStrRev(Left(Prefix, Len(Prefix) - 1), "."));
          }
        } else {
          pK = Prefix + LCase(SplitWord(M, 1, "="));
          pV = ConvertProperty(SplitWord(M, 2, "=", true, true));
          // TODO (not supported): On Error Resume Next
          Props.Add(pV, pK);
          // TODO (not supported): On Error GoTo 0
        }
      } while(!(true);
      K = K + J - 1;
      R = R + sSpace(I * SpIndent) + StartControl(L, Props, LMatch(M, "End"), CodeSection, Tag) + vbCrLf;
      I = I + 1;
      Stck[I] = Tag;
    } else if (L == "End") {
      Props = null;
      Tag = Stck[I];
      I = I - 1;
      if (Tag != "") {
        R = R + sSpace(I * SpIndent) + EndControl(Tag) + vbCrLf;
      }
    }
NextLine:
  }
  ConvertFormUi = R;
  return ConvertFormUi;
}


private static string StartControl(string L, Collection Props, bool DoEmpty, string Code, out string TagType) {
  string StartControl = "";
  string cType = "";
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
  cName = SplitWord(L, 3);
  cIndex = cValP(ref Props, ref "Index");
  if (cIndex != "") {
    cName = cName + "_" + cIndex;
  }

  ControlData(cType, tType, tCont, tDef, Features);

  S = "";
  // TODO (not supported): On Error Resume Next
  if (tType == "Line" || tType == "Shape" || tType == "Timer") {
    return StartControl;

  } else if (tType == "Window") {
    S = S + M + "<Window x:Class=\"" + AssemblyName() + ".Forms." + cName + "\"";
    S = S + N + "    xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"";
    S = S + N + "    xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\"";
    S = S + N + "    xmlns:d=\"http://schemas.microsoft.com/expression/blend/2008\"";
    S = S + N + "    xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"";
    S = S + N + "    xmlns:local=\"clr-namespace:" + AssemblyName() + ".Forms\"";
    S = S + N + "    xmlns:usercontrols=\"clr-namespace:" + AssemblyName() + ".UserControls\"";
    S = S + N + "    mc:Ignorable=\"d\"";
    S = S + N + "    Title=" + Quote(cValP(ref Props, ref "caption"));
    S = S + M + "    Height=" + Quote(Px(cValP(ref Props, ref "clientheight", 0) + 435));
    S = S + M + "    Width=" + Quote(Px(cValP(ref Props, ref "clientwidth", 0) + 435));
    S = S + CheckControlEvents("Window", "Form", Code);
    S = S + M + ">";
    S = S + N + " <Grid";
  } else if (tType == "GroupBox") {
    S = S + "<" + tType;
    S = S + " x:Name=\"" + cName + "\"";

    S = S + " Margin=" + Quote(Px(cValP(ref Props, ref "left")) + "," + Px(cValP(ref Props, ref "top")) + ",0,0");
    S = S + " Width=" + Quote(Px(cValP(ref Props, ref "width")));
    S = S + " Height=" + Quote(Px(cValP(ref Props, ref "height")));
    S = S + " VerticalAlignment=\"Top\"";
    S = S + " HorizontalAlignment=\"Left\"";
    S = S + " FontFamily=" + Quote(cValP(ref Props, ref "font.name", "Calibri"));
    S = S + " FontSize=" + Quote(cValP(ref Props, ref "font.size", 10));

    S = S + " Header=\"" + cValP(ref Props, ref "caption") + "\"";
    S = S + "> <Grid Margin=\"0,-15,0,0\"";
  } else if (tType == "Canvas") {
    S = S + "<" + tType;
    S = S + " x:Name=\"" + cName + "\"";

    S = S + " Margin=" + Quote(Px(cValP(ref Props, ref "left")) + "," + Px(cValP(ref Props, ref "top")) + ",0,0");
    S = S + " Width=" + Quote(Px(cValP(ref Props, ref "width")));
    S = S + " Height=" + Quote(Px(cValP(ref Props, ref "height")));
  } else if (tType == "Image") {
    S = S + "<" + tType;

    S = S + " x:Name=\"" + cName + "\"";
    S = S + " Margin=" + Quote(Px(cValP(ref Props, ref "left")) + "," + Px(cValP(ref Props, ref "top")) + ",0,0");
    S = S + " Width=" + Quote(Px(cValP(ref Props, ref "width")));
    S = S + " Height=" + Quote(Px(cValP(ref Props, ref "height")));
    S = S + " VerticalAlignment=" + Quote("Top");
    S = S + " HorizontalAlignment=" + Quote("Left");
  } else {
    S = "";
    S = S + "<" + tType;
    S = S + " x:Name=\"" + cName + "\"";
    S = S + " Margin=" + Quote(Px(cValP(ref Props, ref "left")) + "," + Px(cValP(ref Props, ref "top")) + ",0,0");
    S = S + " Padding=" + Quote("2,2,2,2");
    S = S + " Width=" + Quote(Px(cValP(ref Props, ref "width")));
    S = S + " Height=" + Quote(Px(cValP(ref Props, ref "height")));
    S = S + " VerticalAlignment=" + Quote("Top");
    S = S + " HorizontalAlignment=" + Quote("Left");

  }

  if (IsInStr(Features, "Font")) {
    S = S + " FontFamily=" + Quote(cValP(ref Props, ref "font.name", "Calibri"));
    S = S + " FontSize=" + Quote(cValP(ref Props, ref "font.size", 10));
    if (Val(cValP(ref Props, ref "font.weight", "400")) > 400) {
      S = S + " FontWeight=" + Quote("Bold");
    }

  }

  if (IsInStr(Features, "Content")) {
    S = S + " Content=" + QuoteXML(cValP(ref Props, ref "caption") + cValP(ref Props, ref "text"));
  }

  if (IsInStr(Features, "Header")) {
    S = S + " Content=" + QuoteXML(cValP(ref Props, ref "caption") + cValP(ref Props, ref "text"));
  }

  V = cValP(ref Props, ref "caption") + cValP(ref Props, ref "text");
  if (IsInStr(Features, "Text") && V != "") {
    S = S + " Text=" + QuoteXML(V);
  }

  V = cValP(ref Props, ref "ToolTipText");
  if (IsInStr(Features, "ToolTip") && V != "") {
    S = S + " ToolTip=" + Quote(V);
  }

  S = S + CheckControlEvents(tType, cName, Code);

  if (DoEmpty) {
    S = S + " />";
    TagType = "";
  } else {
    S = S + ">";
    TagType = tType;
  }
  StartControl = S;
  return StartControl;
}

public static string CheckControlEvents(string ControlType, string ControlName, string CodeSection= "") {
  string CheckControlEvents = "";
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
  if (HasFocus) {
    Res = Res + CheckEvent("GotFocus", ControlName, ControlType, CodeSection);
    Res = Res + CheckEvent("LostFocus", ControlName, ControlType, CodeSection);
    Res = Res + CheckEvent("KeyDown", ControlName, ControlType, CodeSection);
    Res = Res + CheckEvent("KeyUp", ControlName, ControlType, CodeSection);
  }
  if (HasClick) {
    Res = Res + CheckEvent("Click", ControlName, ControlType, CodeSection);
    Res = Res + CheckEvent("DblClick", ControlName, ControlType, CodeSection);
  }
  if (HasChange) {
    Res = Res + CheckEvent("Change", ControlName, ControlType, CodeSection);
  }
  if (IsWindow) {
    Res = Res + CheckEvent("Load", ControlName, ControlType, CodeSection);
    Res = Res + CheckEvent("Unload", ControlName, ControlType, CodeSection);
//    Res = Res & CheckEvent("QueryUnload", ControlName, ControlType, CodeSection)
  }

  CheckControlEvents = Res;
  return CheckControlEvents;
}

public static string CheckEvent(string EventName, string ControlName, string ControlType, string CodeSection= "") {
  string CheckEvent = "";
  string Search = "";
  string Target = "";
  string N = "";

  int L = 0;
  string V = "";

  N = ControlName + "_" + EventName;
  Search = " " + N + "(";
  Target = EventName;
  switch(EventName) {
    case "DblClick":
      Target = "MouseDoubleClick";
      break;
    case "Change":
      if (ControlType == "TextBox") {
        Target = "TextChanged";
      }
      break;
    case "Load":
      Target = "Loaded";
      break;
    case "Unload":
      Target = "Unloaded";
break;
}
  L = InStr(1, CodeSection, Search, vbTextCompare);
  if (L > 0) {
    V = Mid(CodeSection, L + 1, Len(N)); // Get exact capitalization from source....
    CheckEvent = " " + Target + "=\"" + V + "\"";
  } else {
    CheckEvent = "";
  }
  return CheckEvent;
}

public static string EndControl(string tType) {
  string EndControl = "";
  switch(tType) {
    case "Line":
      EndControl = "";
      break;
    case "Window":
      EndControl = " </Grid>" + vbCrLf + "</Window>";
      break;
    case "GroupBox":
      EndControl = "</Grid> </GroupBox>";
      break;
    default:
      EndControl = "</" + tType + ">";
break;
}
  return EndControl;
}

public static bool IsEvent(string Str) {
  bool IsEvent = false;
  IsEvent = EventStub(Str) != "";
  return IsEvent;
}

public static string EventStub(string FName) {
  string EventStub = "";
  string S = "";
  string C = "";
  string K = "";


  C = SplitWord(FName, 1, "_");
  K = SplitWord(FName, 2, "_");
  switch(K) {
    case "Click":
      S = "private void " + FName + "(object sender, RoutedEventArgs e) { " + FName + "(); }" + vbCrLf;
      break;
    case "Change":
      S = "private void " + C + "_Change(object sender, System.Windows.Controls.TextChangedEventArgs e) { " + FName + "(); }" + vbCrLf;
      break;
    case "QueryUnload":
      S = "private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) { int c = 0, u = 0 ;  " + FName + "(out c, ref u); e.Cancel = c != 0;  }" + vbCrLf;
//      V = " long doCancel; long UnloadMode; " & FName & "(ref doCancel, ref UnloadMode);"
      break;
    case "Validate":
//      V = "long doCancel; " & FName & "(ref doCancel);"
      break;
    case "KeyDown":
      break;
    case "MouseMove":
break;
}

  EventStub = S;
  return EventStub;
}
}
