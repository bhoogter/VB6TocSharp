using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
//using static System.Drawing.Printing.PrinterSettings;
using static Microsoft.VisualBasic.Constants;

public static class VBExtension
{
    //private static Printer mPrinter = new Printer();
    //private static List<Printer> mPrinters = null;
    public enum vbTriState { vbFalse = 0, vbTrue = -1, vbUseDefault = -2 }

    public static int MousePointer { get { return 0; } set { } }
    public static int SenderIndex(string name) { return ValI(name.Substring(name.LastIndexOf('_') + 1)); }
    public static int SenderIndex(object sender) { return SenderIndex(((FrameworkElement)sender).Name); }

    //public static List<Printer> Printers
    //{
    //    get
    //    {
    //        if (mPrinters == null)
    //        {
    //            mPrinters = new List<Printer>();
    //            foreach (var P in new PrinterCollection()) mPrinters.Add((Printer)P);
    //        }
    //        return mPrinters;
    //    }
    //}

    //public static List<string> PrinterNames()
    //{
    //    StringCollection col = System.Drawing.Printing.PrinterSettings.InstalledPrinters;
    //    string[] arr = new string[col.Count];
    //    col.CopyTo(arr, 0);
    //    return new List<string>(arr);
    //}
    //public static void ResetPrinters() { mPrinters = null; }
    //public static Printer Printer
    //{
    //    get => mPrinter ?? new Printer();
    //    set { foreach (var P in Printers) { if (((Printer)P).DeviceName == value.DeviceName) mPrinter = P; } }
    //}
    //public static string PrinterName() { return VBExtension.Printer != null ? Printer.DeviceName : ""; }
    //public static void PrintPicture(this Printer P, BitmapImage I, dynamic x1 = null, dynamic y1 = null, dynamic w1 = null, dynamic h1 = null, dynamic x2 = null, dynamic y2 = null, dynamic w2 = null, dynamic h2 = null)
    //{ System.Drawing.Image I2 = null; P.PaintPicture(I2, ValF(x1), ValF(y1), ValF(w1), ValF(h1), ValF(x2), ValF(h2), ValF(w2), ValF(h2)); }
    //public static void PrintPicture(this Printer P, ImageSource I, dynamic x1 = null, dynamic y1 = null, dynamic w1 = null, dynamic h1 = null, dynamic x2 = null, dynamic y2 = null, dynamic w2 = null, dynamic h2 = null)
    //{ System.Drawing.Image I2 = null; P.PaintPicture(I2, ValF(x1), ValF(y1), ValF(w1), ValF(h1), ValF(x2), ValF(h2), ValF(w2), ValF(h2)); }
    //public static void PaintPicture(this Printer P, System.Windows.Controls.Image I, dynamic x1 = null, dynamic y1 = null, dynamic w1 = null, dynamic h1 = null, dynamic x2 = null, dynamic y2 = null, dynamic w2 = null, dynamic h2 = null)
    //{ System.Drawing.Image I2 = null; P.PaintPicture(I2, ValF(x1), ValF(y1), ValF(w1), ValF(h1), ValF(x2), ValF(h2), ValF(w2), ValF(h2)); }
    //public static void PaintPicture(this Image P, System.Windows.Controls.Image I, dynamic x1 = null, dynamic y1 = null, dynamic w1 = null, dynamic h1 = null, dynamic x2 = null, dynamic y2 = null, dynamic w2 = null, dynamic h2 = null)
    //{ System.Drawing.Image I2 = null; }
    //public static void PrintNNL(this Printer P, params String[] s) { float Y = P.CurrentY; P.Print(s); P.CurrentY = Y; }
    //public static void Line(this Printer P, float X1, float Y1, float X2, float Y2, int Style = 0, bool Box = false) { }
    //public static void LineStep(this Printer P, float X1, float Y1, float X2, float Y2, int Style = 0, bool Box = false) { }
    //public static void Box(this Printer P, float X1, float Y1, float X2, float Y2, int Style = 0)
    //{
    //    P.Line(X1, Y1, X2, Y2, Style, true);
    //}
    //public static void BoxStep(this Printer P, float X1, float Y1, float X2, float Y2, int Style = 0)
    //{
    //    P.Line(X1, Y1, X1 + X2, Y1 + Y2, Style, true);
    //}
    //public static void Circle(this Printer P, float X1, float Y1, float X2, float Y2, float Radius = 0, bool Box = false) { }
    public static string Tab(int N) { return "::TABSTOP:" + N; }

    public static void Load(Window Ob) { }
    public static void CenterInScreen(this Window Ob)
    {
        Ob.Left = (System.Windows.SystemParameters.PrimaryScreenWidth - Ob.Width) / 2;
        Ob.Top = (System.Windows.SystemParameters.PrimaryScreenHeight - Ob.Height) / 2;
    }
    //public static void Unload(this Window Ob) { Ob.Close(); }
    public static void Stop(int code = 1) { Environment.Exit(code); }

    public static bool IIf(bool A, bool B, bool C) { return !!A ? B : C; }
    public static string IIf(bool A, string B, string C) { return !!A ? B : C; }
    public static double IIf(bool A, double B, double C) { return !!A ? B : C; }
    public static decimal IIf(bool A, decimal B, decimal C) { return !!A ? B : C; }
    public static int IIf(bool A, int B, int C) { return !!A ? B : C; }

    public static double ValDouble(string s)
    {
        string f = "";

        if (s == null) return 0;
        if (s.Equals("true", StringComparison.OrdinalIgnoreCase)) return 1;
        if (s.Equals("false", StringComparison.OrdinalIgnoreCase)) return 1;

        if (s.StartsWith("-")) { f = "-"; s = s.Substring(1); }
        for (int i = 0; i < s.Length; i++)
        {
            char c = s.Substring(i, 1)[0];
            if (c >= '0' && c <= '9' || c == '.') f += c.ToString();
            else break;
        }
        if (f == "") return 0;
        return double.Parse(f);
    }

    public static decimal ValD(string A) { return (decimal)ValDouble((A ?? "").Replace(",", "")); }
    public static decimal ValD(decimal A) { return A; }
    public static decimal ValD(int A) { return A; }
    public static decimal ValD(double A) { return (decimal)A; }
    public static float ValF(string A) { return ValF(ValD(A)); }
    public static float ValF(decimal A) { return (float)A; }
    public static int ValI(string A) { return (int)ValDouble(A); }
    public static int ValL(string A) { return ValI(A); }
    public static int ValI(int A) { return A; }
    public static int ValI(decimal A) { return (int)A; }
    public static int ValI(float A) { return (int)A; }
    public static int ValI(double A) { return (int)A; }
    public static int ValI(bool A) { return A ? 1 : 0; }
    public static DateTime IIf(bool A, DateTime B, DateTime C) { return !!A ? B : C; }

    public static string Spc(int I) { return Strings.StrDup(I, ' '); }
    public static decimal TextWidth(string S) { return S.Length * 10m; }
    public static decimal TextHeight(string S) { return modTextFiles.CountLines(S) * 10m; }
    public static int ScaleX(int X, dynamic A, dynamic B) { return X; }
    public static int ScaleY(int Y, dynamic A, dynamic B) { return Y; }

    public static bool IsEmpty(object A) { return false; }
    public static bool IsMissing(object A) { return false; }
    public static bool IsNull(object A) { return A == null || (A is System.DBNull); }
    public static bool IsNothing(object A) { return IsNull(A); }
    public static bool IsObject(object A) { return !IsNothing(A); }

    public static bool IsDate(string D) { try { System.DateTime.Parse(D); } catch { return false; } return true; }

    public static DateTime CDate(dynamic A) { if (A is DateTime) return A; return IsDate(A.ToString()) ? DateTime.MinValue : System.DateTime.Parse(A.ToString()); }
    public static double CDbl(object A) { return (A is System.IConvertible) ? ((System.IConvertible)A).ToDouble(null) : 0; }
    public static decimal CDec(object A) { return (decimal)((A is System.IConvertible) ? ((System.IConvertible)A).ToDouble(null) : 0); }
    public static long CLng(object A) { return (A is System.IConvertible) ? ((System.IConvertible)A).ToInt64(null) : 0; }
    public static int CInt(object A) { return (A is System.IConvertible) ? ((System.IConvertible)A).ToInt32(null) : 0; }
    public static short CShort(object A) { short z = 0; return (A is System.IConvertible) ? ((System.IConvertible)A).ToInt16(null) : z; }
    public static string CStr(object A) { return "" + A; }
    public static bool CBool(object A) { { return (A is System.IConvertible) ? ((System.IConvertible)A).ToBoolean(null) : false; } }
    public static decimal CCur(decimal A) { return A; }
    //public static string Chr(int C) { return Chr((int)C); }
    //public static string Mid(string S, int F) { return Mid(S, (int)F); }
    //public static string Mid(string S, int F, int L) { return Mid(S, (int)F, (int)L); }
    //public static string Left(string S, int F) { return Left(S, (int)F); }
    //public static string Right(string S, int F) { return Right(S, (int)F); }
    public static decimal RndD() { return (decimal)VBMath.Rnd(); }

    public static List<Window> Forms
    {
        get
        {
            List<Window> ret = new List<Window>();
            if (Application.Current == null) return ret;
            foreach (Window w in Application.Current.Windows) ret.Add(w);
            return ret;
        }
    }

    public static System.Windows.Media.Brush ColorToBrush(String C)
    { return (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString(C); }
    public static System.Windows.Media.Brush ColorToBrush(uint C)
    { return (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#" + C.ToString("X")); }

    public static DateInterval getDateInterval(string s)
    {
        switch (s)
        {
            case "y": return DateInterval.Year;
            case "m": return DateInterval.Month;
            case "w": return DateInterval.WeekOfYear;
            case "h": return DateInterval.Hour;
            case "d": return DateInterval.Day;
            case "n": return DateInterval.Minute;
            case "s": return DateInterval.Second;
            default: return DateInterval.Day;
        }
    }
    //public static DateTime DateAdd1(string unit, int amount, DateTime when) { return DateAndTime.DateAdd(getDateInterval(unit), amount, when); }

    public static System.DateTime DateValue(object A)
    {
        if (A is String) try { return System.DateTime.Parse((string)A); } catch { return DateTime.MinValue; }
        return CDate(A);
    }

    public static bool IsList(object A) { return A != null && (A is System.Collections.IList); }
    public static int LBound(object A) { return A != null && (A is System.Collections.IList) ? (((System.Collections.IList)A).Count == 0 ? -1 : 0) : 0; }
    public static int UBound(object A) { return A != null && (A is System.Collections.IList) ? ((System.Collections.IList)A).Count - 1 : 0; }

    public static bool IsLike(string A, string B) { return Microsoft.VisualBasic.CompilerServices.LikeOperator.LikeString(A, B, Microsoft.VisualBasic.CompareMethod.Binary); }

    public static bool VBOpenFile(dynamic A, dynamic B) { return false; }
    public static bool VBWriteFile(dynamic A, dynamic B) { return false; }
    public static bool VBCloseFile(dynamic A) { return false; }
    public static string VBReadFileLine(dynamic A, dynamic B) { return ""; }
    public static bool DoEvents(Window Frm = null)
    {
        //if (Frm == null) Frm = MainMenu1.instance;
        Frm.Dispatcher.Invoke(new Action(delegate () { }), DispatcherPriority.ContextIdle);
        return true;
    }

    public static bool Resume() { return false; }
    public static bool End() { Application.Current.Shutdown(); return false; }

    public static ModifierKeys Shift { get { return Keyboard.Modifiers; } }
    public static string ShiftStr(Key v)
    { return ShiftStr(v.ToString()); }
    public static string ShiftStr(string v = null)
    {
        string s = "";
        if (0 != (Keyboard.Modifiers & ModifierKeys.Windows)) s += "Win-";
        if (0 != (Keyboard.Modifiers & ModifierKeys.Control)) s += "Ctrl-";
        if (0 != (Keyboard.Modifiers & ModifierKeys.Alt)) s += "Alt-";
        if (0 != (Keyboard.Modifiers & ModifierKeys.Shift)) s += "Shift-";
        if (v != null) s += v.ToString();
        return s;
    }

    public static bool HasEmptyText(this TextBox textBox) { return string.IsNullOrEmpty(textBox.Text); }
    public static decimal getValue(this TextBox textBox) { try { return Decimal.Parse(textBox.Text); } catch { return 0; } }
    //public static decimal setValue(this TextBox textBox, decimal value) { textBox.Text = FormatQuantity(value); return getValue(textBox); }
    //public static decimal setCurrency(this TextBox c, decimal value) { c.Text = CurrencyFormat(value); return c.getCurrency(); }
    public static decimal getCurrency(this TextBox c) { return ValD(c.Text); }
    public static decimal getValueCurrency(this TextBox c) { return ValD(c.Text); }
    //public static decimal setValueCurrency(this TextBox c, decimal value) { c.Text = CurrencyFormat(value); return c.getCurrency(); }
    public static int getValueLong(this TextBox textBox) { try { return int.Parse(textBox.Text); } catch { return 0; } }
    public static int setValueLong(this TextBox textBox, int value) { textBox.Text = value.ToString(); return getValueLong(textBox); }
    public static DateTime? getValueDate(this TextBox textBox, DateTime? defaultDate = null) { try { return DateValue(textBox.Text); } catch { return defaultDate; } }
    public static DateTime? setValueDate(this TextBox textBox, DateTime? value) { textBox.Text = value == null ? "" : ((DateTime)value).ToShortDateString(); return textBox.getValueDate(); }
    public static decimal getValue(this Label label) { try { return ValD(label.Content.ToString()); } catch { return 0; } }
    //public static decimal setValue(this Label label, decimal value) { label.Content = FormatQuantity(value); return getValue(label); }
    //public static decimal setCurrency(this Label c, decimal value) { c.Content = CurrencyFormat(value); return c.getCurrency(); }
    public static decimal getCurrency(this Label c) { return ValD(c.Content.ToString()); }
    public static decimal getValueCurrency(this Label c) { return ValD(c.Content.ToString()); }
    //public static decimal setValueCurrency(this Label c, decimal value) { c.Content = CurrencyFormat(value); return c.getCurrency(); }
    public static int getValueLong(this Label textBox) { try { return int.Parse(textBox.Content.ToString()); } catch { return 0; } }
    public static int setValueLong(this Label textBox, int value) { textBox.Content = value.ToString(); return getValueLong(textBox); }
    public static DateTime? getValueDate(this Label textBox, DateTime? defaultDate = null) { try { return DateValue(textBox.Content.ToString()); } catch { return defaultDate; } }
    public static DateTime? setValueDate(this Label textBox, DateTime? value) { textBox.Content = value == null ? "" : ((DateTime)value).ToShortDateString(); return textBox.getValueDate(); }

    public static bool getValue(this CheckBox chk) { try { return ((bool)chk.IsChecked); } catch { return false; } }
    public static bool setValue(this CheckBox chk, bool value) { chk.IsChecked = value; return getValue(chk); }
    //    public static int getValue(this CheckBox chk) { try { return ((bool)chk.IsChecked); } catch { return false; } }
    public static int setValue(this CheckBox chk, int value) { chk.IsChecked = value != 1; return getValue(chk) ? 1 : 0; }

    public static bool getValue(this Button btn) { try { return ((bool)btn.IsPressed); } catch { return false; } }
    public static bool setValue(this Button btn, bool value) { try { btn.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); return true; } catch { return false; } }
    public static BitmapImage getImage(this Button btn)
    {
        try
        {
            Image T = null;
            dynamic c = btn.Content;

            if (c is Image) T = c;
            if (c is Panel)
                foreach (var l in c.Children) if (l is Image) { T = (Image)l; break; }

            if (T is null) return null;
            return (BitmapImage)T.Source;
        }
        catch { return null; }
    }
    public static BitmapImage setImage(this Button Cmd, BitmapImage value)
    {
        try
        {
            if (Cmd.Content is string)
            {
                string caption = Cmd.Content.ToString();
                Canvas C = new Canvas();
                Cmd.Content = C;
                C.Width = Cmd.Width;
                C.Height = Cmd.Height;
                Label L = new Label();
                L.Content = caption;
                C.Children.Add(L);
                L.FontSize = 12d;
                L.Padding = new Thickness(0);
                L.Width = L.MeasureString(caption).Width;
                L.Height = L.MeasureString(caption).Height;
                Canvas.SetLeft(L, (Cmd.Width - L.Width) / 2);
                Canvas.SetTop(L, Cmd.Height - L.Height - 10);
                Image I = new Image();
                C.Children.Add(I);
                I.Width = Cmd.Width - 10;
                I.Height = Cmd.Height - L.Height - 12;
                I.Stretch = System.Windows.Media.Stretch.Uniform;
                I.Source = value;
                Canvas.SetLeft(I, (Cmd.Width - I.Width) / 2);
                Canvas.SetTop(I, 0);
                return value;
            }
            else if (Cmd.Content is Panel)
            {
                Image I = (Image)((Panel)Cmd.Content).ControlOf(typeof(Image), 0);
                if (I == null) return null;
                I.Source = value;
                return value;
            }
            else if (Cmd.Content is Image)
            {
                ((Image)Cmd.Content).Source = value;
                return value;
            }
        }
        catch { return null; }
        return null;
    }
    public static string getCaption(this Button btn)
    {
        try
        {
            Label T = null;
            foreach (var c in ((Panel)btn.Content).Children) if (c is Label) { T = (Label)c; break; }
            if (T is null) return "";
            if (T.Content is null) return "";
            return T.Content.ToString();
        }
        catch { return ""; }
    }
    public static string setCaption(this Button btn, string value)
    {
        Label T = null;
        if (btn.Content is Panel)
            foreach (var c in ((Panel)btn.Content).Children) if (c is Label) { T = (Label)c; break; }
        if (btn.Content is Label) T = (Label)btn.Content;
        if (btn.Content is string) btn.Content = value.Replace("&", "_");
        if (T is null) return "";
        return (string)(T.Content = value.Replace("&", "_"));
    }

    public static bool getVisible(this FrameworkElement c) { return c.Visibility == System.Windows.Visibility.Visible; }
    public static bool isVisible(this FrameworkElement c) { if (c == null) return false; return c.Visibility == System.Windows.Visibility.Visible; }
    public static bool setVisible(this FrameworkElement c, bool value, bool CollapseClose = false) { if (c == null) return false; c.Visibility = value ? Visibility.Visible : (CollapseClose ? Visibility.Collapsed : Visibility.Hidden); return c.getVisible(); }
    public static bool Locked(this TextBox t, bool value = true) { return false; }
    public static bool Locked(this ComboBox t, bool value = true) { return false; }
    public static bool Locked(this ListBox t, bool value = true) { return false; }
    public static bool setToolTipText(this FrameworkElement c, string Id) { return true; }
    public static string getToolTipText(this FrameworkElement c) { return ""; }
    public static void ZOrder(this FrameworkElement c, int v) { }
    public static void SelectContents(this TextBox c) { c.SelectionStart = 0; c.SelectionLength = c.Text.Length; }
    public static void FocusSelect(this TextBox c) { c.SelectionStart = 0; c.SelectionLength = c.Text.Length; c.Focus(); }
    public static void SelectContents(this ComboBox c) { }
    public static void toUpper(this TextBox c) { if (c.Text != c.Text.ToUpper()) c.Text = c.Text.ToUpper(); }

    public static string getText(this RichTextBox r) { return ""; }
    public static string setText(this RichTextBox r, string v) { return ""; }
    public static bool LoadFile(this RichTextBox r, string f) { return true; }

    private static Action EmptyDelegate = delegate () { };

    public static void Refresh(this UIElement uiElement) { uiElement.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate); }
    public static bool SetFocus(this FrameworkElement c) { try { return c.Focus(); } catch { return false; } }
    //public static bool Move(this Control c, double X = -10000, double Y = -10000, double W = -1000, double H = -10000, bool MakeVisible = false)
    //{ return c.Move((decimal)X, (decimal)Y, (decimal)W, (decimal)H, MakeVisible); }
    public static bool Move(this FrameworkElement c, decimal X = -10000, decimal Y = -10000, decimal W = -10000, decimal H = -10000, bool MakeVisible = false)
    {
        if (c == null) return false;
        if (W > 0) c.Width = (double)W;
        if (H > 0) c.Height = (double)H;
        //if (c.Parent is Grid) {
        Thickness t = c.Margin;
        if (X != -10000 && X != -1) t.Left = (double)X;
        if (Y != -10000 && Y != -1) t.Top = (double)Y;
        c.Margin = t;
        //}
        //else if (c.Parent is Canvas)
        //{
        //    if (X != -10000 && X != -1) Canvas.SetLeft(c, (double)X);
        //    if (Y != -10000 && Y != -1) Canvas.SetTop(c, (double)Y);
        //}

        c.Margin = new Thickness(
            X == -10000 || X == -1 ? c.Margin.Left : (double)X,
            Y == -10000 || Y == -1 ? c.Margin.Top : (double)Y,
            0, 0
            );
        if (MakeVisible) c.Visibility = Visibility.Visible;
        //try { c.Focus(); } catch { }
        return false;
    }
    public static bool Move(this FrameworkElement c, double X = -10000, double Y = -10000, double W = -10000, double H = -10000, bool MakeVisible = false)
    { return c.Move((decimal)X, (decimal)Y, (decimal)W, (decimal)H, MakeVisible); }
    public static void setWindowState(this Window w, WindowState X) { w.WindowState = X; }

    public static string AppHelpFile() { return ""; }
    public static bool Show(this Window w, int Modal) { w.ShowDialog(); return true; }
    //public static bool Load(this Window w) { return true; }
    public static void setHelpContextID(this Window w, int Id) { }
    public static int getHelpContextID(this Window wId) { return 0; }
    public static double ScaleWidth(this Window w) { return w.Width; ; }
    public static double ScaleHeight(this Window w) { return w.Height; }
    public static IntPtr hWnd(this Window w) { return new WindowInteropHelper(Window.GetWindow(w)).Handle; }
    public static IntPtr hWnd(this FrameworkElement w) { return new WindowInteropHelper(Window.GetWindow(w)).Handle; }
    public static bool getVisible(this Window w) { return w.Visibility == System.Windows.Visibility.Visible; }
    public static bool setVisible(this Window w, bool value) { w.Visibility = value ? System.Windows.Visibility.Visible : System.Windows.Visibility.Hidden; return w.getVisible(); }


    public static string SanitizeNls(string s)
    {
        return SanitizeNls(s, vbLf);
    }

    public static string SanitizeNls(string s, string desired)
    {
        bool hCr = s.Contains(vbCr);
        bool hLf = s.Contains(vbLf);
        if (hCr && !hLf) return s.Replace(vbCr, desired);
        if (hLf && !hCr) return s.Replace(vbLf, desired);
        return s.Replace(vbCrLf, desired);
    }

    public static Size MeasureString(this Label el, string candidate)
    {
        var formattedText = new FormattedText(SanitizeNls(candidate, "\n"), CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            new Typeface(el.FontFamily, el.FontStyle, el.FontWeight, el.FontStretch),
            el.FontSize, Brushes.Black, new NumberSubstitution(), TextFormattingMode.Display);
        return new Size(formattedText.Width, formattedText.Height);
    }
    public static Size MeasureString(this TextBox el, string candidate)
    {
        var formattedText = new FormattedText(candidate, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            new Typeface(el.FontFamily, el.FontStyle, el.FontWeight, el.FontStretch),
            el.FontSize, Brushes.Black, new NumberSubstitution(), TextFormattingMode.Display);
        return new Size(formattedText.Width, formattedText.Height);
    }
    public static Size MeasureString(this Window el, string candidate)
    {
        var formattedText = new FormattedText(candidate, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            new Typeface(el.FontFamily, el.FontStyle, el.FontWeight, el.FontStretch),
            el.FontSize, Brushes.Black, new NumberSubstitution(), TextFormattingMode.Display);
        return new Size(formattedText.Width, formattedText.Height);
    }
    public static double TextWidth(this Canvas t, string s) { return ((Window)t.Parent).MeasureString(s).Width; }
    public static double TextHeight(this Canvas t, string s) { return ((Window)t.Parent).MeasureString(s).Height; }
    public static double TextWidth(this Label t, string s) { return t.MeasureString(s).Width; }
    public static double TextHeight(this Label t, string s) { return t.MeasureString(s).Height; }
    public class ScreenMetrics
    {
        public int Width => (int)System.Windows.SystemParameters.PrimaryScreenWidth;
        public int Height => (int)System.Windows.SystemParameters.PrimaryScreenHeight;
        public FrameworkElement ActiveControl;
    }
    public static ScreenMetrics Screen { get => new ScreenMetrics(); }

    public static int itemData(this ComboBox c, int I) { try { return (((ComboboxItem)c.Items[I]).Value); } catch (Exception e) { return 0; } }
    public static int AddItem(this ComboBox c, string C) { return c.Items.Add(new ComboboxItem(C)); }
    public static int AddItem(this ComboBox c, string C, int D) { return c.Items.Add(new ComboboxItem(C, D)); }
    public static int AddItem(this ComboBox c, string C, bool Select) { ComboboxItem x = new ComboboxItem(C); int res = c.Items.Add(x); if (Select) c.SelectedItem = x; return res; }
    public static int AddItem(this ComboBox c, string C, int D, bool Select) { ComboboxItem x = new ComboboxItem(C, D); int res = c.Items.Add(x); if (Select) c.SelectedItem = x; return res; }
    public static String List(this ComboBox c, int Index) { return Index < c.Items.Count ? c.Items[Index].ToString() : null; }
    public static string SetItemText(this ComboBox c, int Index, string Text) { return ((ComboboxItem)c.Items[Index]).Text = Text; }
    public static int SelectedValue(this ComboBox c) { return ((ComboboxItem)c.SelectedItem).Value; }
    public static string SelectedText(this ComboBox c) { return c.SelectedItem == null ? "" : ((ComboboxItem)c.SelectedItem).Text; }
    public static bool SelectText(this ComboBox c, string S) { for (int i = 0; i < c.Items.Count; i++) if (i.ToString() == S) { c.SelectedIndex = i; return true; } return false; }
    public static void RemoveItem(this ComboBox c, int Index) { c.Items.RemoveAt(Index); }
    public static void Clear(this ComboBox c) { c.Items.Clear(); }

    public static int itemData(this ListBox c, int I) { try { return ((int)((ComboboxItem)c.Items[I]).Value); } catch (Exception e) { return 0; } }
    public static bool SelectText(this ListBox c, string S) { for (int i = 0; i < c.Items.Count; i++) if (i.ToString() == S) { c.SelectedIndex = i; return true; } return false; }
    public static int SelectItem(this ListBox c, int I, bool isSelected)
    {
        if (c.SelectionMode == SelectionMode.Multiple)
        { if (isSelected) c.SelectedItems.Add(c.Items[I]); else c.SelectedItems.Remove(c.Items[I]); }
        else { if (isSelected) c.SelectedItem = c.Items[I]; else { if (c.SelectedItem == c.Items[I]) c.SelectedItem = null; } }
        return I;
    }
    public static bool Selected(this ListBox c, int I) { return c.SelectedItems.Contains(c.Items[I]); }
    public static bool Selected(this ListBox c, int I, bool Value)
    {
        if (c.SelectionMode == SelectionMode.Single)
        {
            c.SelectedItem = c.Items[I];
            return true;
        }
        else
        {
            if (Value) c.SelectedItems.Add(c.Items[I]); else c.SelectedItems.Remove(c.Items[I]);
            return c.Selected(I);
        }
    }

    public static int IndexOf(this ListBox c, string s)
    {
        foreach (var l in c.Items) if (Strings.Trim(((ComboboxItem)l).ToString()) == s) return c.Items.IndexOf(l);
        return -1;
    }

    public static bool Contains(this ListBox c, string s) { return c.IndexOf(s) != -1; }

    public static string SelectedText(this ListBox c) { return c.SelectedItem == null ? "" : ((ComboboxItem)c.SelectedItem).ToString(); }
    public static int AddItem(this ListBox c, string C) { return c.Items.Add(new ComboboxItem(C)); }
    public static int AddItem(this ListBox c, string C, int D) { return c.Items.Add(new ComboboxItem(C, D)); }
    public static int AddItem(this ListBox c, string C, bool Selected) { int x = c.Items.Add(new ComboboxItem(C)); return SelectItem(c, x, Selected); }
    public static int AddItem(this ListBox c, string C, int D, bool Selected) { int x = c.Items.Add(new ComboboxItem(C, D)); return SelectItem(c, x, Selected); }
    public static void RemoveItem(this ListBox c, int Index) { c.Items.RemoveAt(Index); }
    //public static string List(this ListBox c, int Index) { return modNumbers.InRange(0, Index, c.Items.Count) ? c.Items[Index].ToString() : ""; }

    public static bool getSelected(this ListBox c, int I) { return c.SelectedItems.Contains(c.Items[I]); }
    public static void setSelected(this ListBox c, int I, bool V)
    {
        if (c.SelectionMode == SelectionMode.Multiple)
        {
            if (V) c.SelectedItems.Add(c.Items[I]);
            else c.SelectedItems.Remove(c.Items[I]);
        }
        else
            c.SelectedItem = c.Items[I];
    }
    public static bool Clear(this ListBox c) { c.Items.Clear(); return true; }

    public static TreeViewItem TreeViewAddItem(TreeView t, string Value, string Key = null, TreeViewItem parent = null)
    {
        TreeViewItem tvi;
        if (parent == null)
        {
            int x = t.Items.Add(new TreeViewItemObject(Value, Key));
            tvi = t.Item(x);
        }
        else
        {
            int x = parent.Items.Add(new TreeViewItemObject(Value, Key));
            tvi = t.Item(x);
        }

        return tvi;

    }
    public static TreeViewItem AddItem(this TreeView t, string Value) { return TreeViewAddItem(t, Value); }
    public static TreeViewItem AddItem(this TreeView t, string Value, string Key) { return TreeViewAddItem(t, Value, Key); }
    public static TreeViewItem AddItem(this TreeView t, string Value, string Key, TreeViewItem parent) { return TreeViewAddItem(t, Value, Key, parent); }
    public static TreeViewItemObject Item(this TreeView t, int x) { return (TreeViewItemObject)t.Items.GetItemAt(x); }
    public static TreeViewItemObject Item(this TreeView t, string x) { foreach (var l in t.Items) if (((TreeViewItemObject)l).getKey() == x) return (TreeViewItemObject)l; return null; }
    public static void setItemColor(this TreeView t, int Item, Brush backColor = null, Brush foreColor = null)
    {
        {
            var actualItem = t.Item(Item);
            if (actualItem != null)
            {
                if (backColor != null) actualItem.Background = backColor;
                if (foreColor != null) actualItem.Foreground = foreColor;
            }
        }
    }
    public static void SetSelectedIndex(this TreeView t, int Index) { ((TreeViewItem)t.Items.GetItemAt(Index)).IsSelected = true; }
    public static void Clear(this TreeView t, string Value, String Key = null) { t.Items.Clear(); }

    public static DateTime getDateTime(this DatePicker DP) { return DP.SelectedDate ?? DP.DisplayDate; }
    public static string getDateString(this DatePicker DP) { return (DP.SelectedDate ?? DP.DisplayDate).ToShortDateString(); }
    public static string getTimeString(this DatePicker DP) { return (DP.SelectedDate ?? DP.DisplayDate).ToShortTimeString(); }

    public static BitmapImage PackageImage(string s, bool placeholder = true)
    {
        if (Strings.Left(s, 1) != "/") s = "/Resources/Images/" + s;
        s = "pack://application:,,," + s;
        try { return new BitmapImage(new Uri(@s)); }
        catch (Exception e)
        {
            if (!placeholder) return null;
            string d = "/Resources/Images/none.bmp";
            return new BitmapImage(new Uri(d, UriKind.Relative));
        }
    }

    public class ComboboxItem
    {
        public ComboboxItem(string vText) { Text = vText; }
        public ComboboxItem(string vText, int vValue) { Text = vText; Value = vValue; }
        public string Text { get; set; }
        public int Value { get; set; }
        public override string ToString() { return Text; }
    }

    public class TreeViewItemObject : TreeViewItem
    {
        string Text;
        string Key;

        public TreeViewItemObject(string Text = "", string Key = "")
        {
            this.Text = Text;
            Header = Text;
            this.Key = Key;
        }

        public new string ToString() { return Text; }
        public string getValue() { return Key; }
        public void setValue(string s) { Key = s; }
        public string getKey() { return Key; }
        public void setKey(string s) { Key = s; }

        public TreeViewItem getContainer(TreeView tv) { return tv.ItemContainerGenerator.ContainerFromItem(this) as TreeViewItem; }
    }

    public class PropIndexer<I, V>
    {
        public delegate void setProperty(I idx, V value);
        public delegate V getProperty(I idx);

        public event getProperty getter;
        public event setProperty setter;

        public PropIndexer(getProperty g, setProperty s) { getter = g; setter = s; }
        public PropIndexer(getProperty g) { getter = g; setter = setPropertyNoop; }
        public PropIndexer() { getter = getPropertyNoop; setter = setPropertyNoop; }

        public void setPropertyNoop(I idx, V value) { }
        public V getPropertyNoop(I idx) { return default(V); }

        public V this[I idx]
        {
            get => getter.Invoke(idx);
            set => setter.Invoke(idx, value);
        }
    }
    public class PropIndexer2<I, J, V>
    {
        public delegate void setProperty(I idx, J idx2, V value);
        public delegate V getProperty(I idx, J idx2);

        public event getProperty getter;
        public event setProperty setter;

        public PropIndexer2(getProperty g, setProperty s) { getter = g; setter = s; }
        public PropIndexer2(getProperty g) { getter = g; setter = setPropertyNoop; }
        public PropIndexer2() { getter = getPropertyNoop; setter = setPropertyNoop; }

        public void setPropertyNoop(I idx, J idx2, V value) { }
        public V getPropertyNoop(I idx, J idx2) { return default(V); }

        public V this[I idx, J idx2]
        {
            get => getter.Invoke(idx, idx2);
            set => setter.Invoke(idx, idx2, value);
        }
    }

    public static dynamic CreateObject(string IdName)
    {
        Type ObjectType = Type.GetTypeFromProgID(IdName);
        dynamic ObjectInst = Activator.CreateInstance(ObjectType);
        return ObjectInst;
    }
    public class Timer
    {
        private System.Windows.Threading.DispatcherTimer tmr = new System.Windows.Threading.DispatcherTimer();
        public Action Action;
        private void dispatcherTimer_Tick(object sender, EventArgs e) { if (Action != null) Action.Invoke(); }

        public Timer(Action e = null, int vInterval = 1000, bool vEnabled = false)
        {
            tmr.Tick += dispatcherTimer_Tick;
            Action = e;
            Interval = vInterval;
            Enabled = vEnabled;
        }

        public System.Windows.Threading.DispatcherTimer timer { get => tmr; }

        public bool IsEnabled
        {
            get => tmr.IsEnabled;
            set { tmr.IsEnabled = value; if (value) tmr.Start(); else tmr.Stop(); }
        }
        public bool Enabled { get => IsEnabled; set => IsEnabled = value; }
        public Timer Discard() { Enabled = false; return null; }

        public int Interval { get => (int)tmr.Interval.TotalMilliseconds; set => tmr.Interval = new TimeSpan(0, 0, 0, 0, value); }
        public int IntervalSeconds { get => (int)tmr.Interval.TotalSeconds; set => tmr.Interval = new TimeSpan(0, 0, 0, value); }
        public dynamic Tag { get; set; }

        public TimeSpan getInterval() { return tmr.Interval; }
        public void setInterval(TimeSpan value) { tmr.Interval = value; }

        public void startTimer(int MilliSeconds) { Enabled = false; Interval = MilliSeconds; Enabled = true; }
        public void startTimerSeconds(int Seconds) { Enabled = false; Interval = Seconds; Enabled = true; }
        public void startTimer(int MilliSeconds, dynamic setTag) { Tag = setTag; startTimer(MilliSeconds); }
        public void startTimerSeconds(int Seconds, dynamic setTag) { Tag = setTag; startTimerSeconds(Seconds); }
        public void stopTimer() { Enabled = false; }
    }
    public static List<FrameworkElement> controlArray(this Window Frm, string name)
    {
        List<FrameworkElement> res = new List<FrameworkElement>();
        Panel G = (Panel)Frm.Content;
        foreach (var C in G.Children)
            if (((FrameworkElement)C).Name.StartsWith(name + "_")) res.Add((FrameworkElement)C);
        return res;
    }
    public static int controlIndex(String name) { try { return ValI(Strings.Mid(name, name.LastIndexOf('_') + 1)); } catch (Exception e) { } return -1; }
    public static int controlIndex(this Control C) { try { return ValI(Strings.Mid(C.Name, C.Name.LastIndexOf('_') + 1)); } catch (Exception e) { } return -1; }
    public static FrameworkElement getControlByIndex(this Window Frm, string Name, int Idx)
    { foreach (var C in Frm.Controls(true)) if (C.Name == Name + "_" + Idx) return C; return null; }
    public static FrameworkElement loadControlByIndex(this Window Frm, Type type, string Name, int Idx = -1)
    {
        FrameworkElement X = Frm.getControlByIndex(Name, Idx);
        if (X != null) return X;
        FrameworkElement C = (FrameworkElement)Activator.CreateInstance(type);
        C.Name = Name + "_" + Idx;
        List<FrameworkElement> els = controlArray(Frm, Name);
        Panel G;
        FrameworkElement el0 = getControlByIndex(Frm, Name, 0);
        if (els.Count > 0) G = els[0].Parent as Panel;
        else if (el0 != null) G = el0.Parent as Panel;
        else G = Frm.Content as Panel;
        G.Children.Add(C);
        return C;
    }
    public static void unloadControlByIndex(this Window Frm, string Name, int Idx = -1)
    {
        FrameworkElement X = Frm.getControlByIndex(Name, Idx);
        if (X != null)
        {
            Panel G = (Panel)Frm.Content;
            G.Children.Remove(X);
        }
    }
    public static void unloadControls(this Window Frm, string Name, int baseIndex = -1)
    {
        Panel G = (Panel)Frm.Content;
        foreach (var C in Frm.Controls())
        {
            string N = ((FrameworkElement)C).Name;
            if (N.StartsWith(Name + "_"))
            {
                if (controlIndex(N) == baseIndex) continue;
                G.Children.Remove(C);
            }
        }
    }
    public static int controlUBound(this Window Frm, string Name)
    {
        int Max = -1;
        foreach (var C in Frm.Controls(true))
        {
            string N = ((FrameworkElement)C).Name;
            if (N.StartsWith(Name + "_"))
            {
                int K = ValI(Strings.Mid(N, N.LastIndexOf('_') + 2));
                if (K > Max) Max = K;
            }
        }
        return Max;
    }

    public static List<FrameworkElement> Controls(this Window w, bool recurse = true)
    {
        Panel g = (Panel)w.Content;
        UIElementCollection children = g.Children;
        List<FrameworkElement> cts = new List<FrameworkElement>();
        foreach (var e in children)
        {
            cts.Add((FrameworkElement)e);
            if (recurse && e is GroupBox)
                foreach (var f in ((GroupBox)e).Controls(recurse)) cts.Add((FrameworkElement)f);
        }
        return cts;
    }
    public static List<FrameworkElement> Controls(this GroupBox w, bool recurse = true)
    {
        Panel g = (Panel)w.Content;
        UIElementCollection children = g.Children;
        List<FrameworkElement> cts = new List<FrameworkElement>();
        foreach (var e in children)
        {
            cts.Add((FrameworkElement)e);
            if (recurse && e is GroupBox)
                foreach (var f in ((GroupBox)e).Controls(recurse)) cts.Add((FrameworkElement)f);
        }
        return cts;
    }

    public static List<string> ControlNames(this Window w, bool recurse = true)
    {
        List<string> res = new List<string>();
        foreach (var c in w.Controls(recurse)) res.Add(c.Name);
        return res;
    }
    public static List<FrameworkElement> Controls(this Window w, Type T)
    {
        List<FrameworkElement> lst = w.Controls(), res = new List<FrameworkElement>();
        foreach (var l in lst) if (l.GetType() == T) res.Add(l);
        return res;
    }
    public static FrameworkElement ControlOf(this Window w, Type T, int n = 0)
    {
        List<FrameworkElement> lst = w.Controls(T);
        if (lst.Count == 0) return null;
        return lst[n < 0 ? 0 : n > lst.Count - 1 ? lst.Count - 1 : n];
    }
    public static FrameworkElement ControlOf(this Panel w, Type T, int n = 0)
    {
        List<FrameworkElement> lst = new List<FrameworkElement>();
        foreach (var l in w.getControls(true)) if (l.GetType() == T) lst.Add(l);
        if (lst.Count == 0) return null;
        return lst[n < 0 ? 0 : n > lst.Count - 1 ? lst.Count - 1 : n];
    }


    public static IEnumerable<FrameworkElement> getControls(this Visual parent, bool recurse = true)
    {
        List<FrameworkElement> res = new List<FrameworkElement>();
        foreach (var el in parent.GetChildren(recurse))
            res.Add((FrameworkElement)el);
        return res;
    }
    public static IEnumerable<Visual> GetChildren(this Visual parent, bool recurse = true)
    {
        if (parent != null)
        {
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                // Retrieve child visual at specified index value.
                var child = VisualTreeHelper.GetChild(parent, i) as Visual;

                if (child != null)
                {
                    yield return child;

                    if (recurse)
                    {
                        foreach (var grandChild in child.GetChildren(true))
                        {
                            yield return grandChild;
                        }
                    }
                }
            }
        }
    }

    //public class KeyedTreeViewItem
    //{
    //    public ObservableCollection<KeyedTreeViewItem> Items { get; set; }
    //    public string Key;
    //    public string Name;
    //    public KeyedTreeViewItem Parent;
    //    private void setup(KeyedTreeViewItem parent, string vKey, string vName)
    //    {
    //        Parent = parent;
    //        Items = new ObservableCollection<KeyedTreeViewItem>();
    //        Key = vKey;
    //        Name = vName;
    //    }

    //    public KeyedTreeViewItem(string vKey, string vName) : base()
    //    { setup(null, vKey, vName); }

    //    private KeyedTreeViewItem(KeyedTreeViewItem parent, string vKey, string vName) : base()
    //    { setup(parent, vKey, vName); }

    //    public void Add(string vKey, string vName)
    //    { Items.Add(new KeyedTreeViewItem(this, vKey, vName)); }

    //    public new string ToString() { return Name; }
    //}
    //public static KeyedTreeViewItem SelectedItemKeyed(this TreeView T)
    //{ return (KeyedTreeViewItem)T.SelectedItem; }

    //public static KeyedTreeViewItem getItemByKey(this TreeView T, string key)
    //{
    //    foreach (KeyedTreeViewItem I in T.Items)
    //        if (I.Key == key) return I;
    //    return null;
    //}


    public static T GetVisualChild<T>(Visual parent) where T : Visual
    {
        T child = default(T);
        int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < numVisuals; i++)
        {
            Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
            child = v as T;
            if (child == null) child = GetVisualChild<T>(v);
            if (child != null) break;
        }
        return child;
    }

    public static DataGridRow GetSelectedRow(this DataGrid grid)
    { return (DataGridRow)grid.ItemContainerGenerator.ContainerFromItem(grid.SelectedItem); }
    public static DataGridRow GetRow(this DataGrid grid, int index)
    {
        DataGridRow row = (DataGridRow)grid.ItemContainerGenerator.ContainerFromIndex(index);
        if (row == null)
        {
            // May be virtualized, bring into view and try again.
            grid.UpdateLayout();
            try { grid.ScrollIntoView(grid.Items[index]); } catch (Exception e) { }
            row = (DataGridRow)grid.ItemContainerGenerator.ContainerFromIndex(index);
        }
        return row;
    }

    public static DataGridCell GetCell(this DataGrid grid, DataGridRow row, int column)
    {
        if (row != null)
        {
            DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(row);

            if (presenter == null)
            {
                grid.ScrollIntoView(row, grid.Columns[column]);
                presenter = GetVisualChild<DataGridCellsPresenter>(row);
            }

            if (presenter == null) return null;
            DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(column);
            return cell;
        }
        return null;
    }

    public static DataGridCell GetCell(this DataGrid grid, int row, int column)
    {
        DataGridRow rowContainer = grid.GetRow(row);
        return grid.GetCell(rowContainer, column);
    }

    public class CommandBase : ICommand
    {
        public event EventHandler CanExecuteChanged;

        private Func<bool> mCanExecute = null;
        private Action<object> mExecute = null;

        public CommandBase(Action<object> vExecute, Func<bool> fCanExecute = null) { mCanExecute = fCanExecute; mExecute = vExecute; }
        public bool CanExecute(object parameter) { return mCanExecute == null ? true : mCanExecute.Invoke(); }
        public void Execute(object parameter) { mExecute.Invoke(parameter); }
    }

    public static dynamic VBSwitch(params dynamic[] vals)
    {
        for (int i = 0; i < vals.Length; i += 2)
        {
            if (i == vals.Length - 1) return vals[i]; // odd number, return as default.
            if (CBool(vals[i])) return vals[i + 1];
        }
        return null;
    }

    // Concatenating an enum yields the NAME of the enum.  This method provides the underlying value (which could be int, long, byte, etc...)
    public static dynamic Value(this Enum item) => Convert.ChangeType(item, item.GetTypeCode());
    public static T Value<T>(this Enum item) => (T)item.Value();
}
