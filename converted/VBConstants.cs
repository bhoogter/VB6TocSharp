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

public static class VBConstants {
  public const long vbKeyLButton = 1; // Left mouse button
    public const long vbKeyRButton = 2;  // CANCEL mouse button 
    public const long vbKeyCancel = 3;  // Middle key  
    public const long vbKeyMButton = 4;  // BACKSPACE mouse button 
    public const long vbKeyBack = 8;  // TAB key  
    public const long vbKeyTab = 9;  //  key  
    public const long vbKeyClear = 12;  //  CLEAR key 
    public const long vbKeyReturn = 13;  //  ENTER key 
    public const long vbKeyShift = 16;  //  SHIFT key 
    public const long vbKeyControl = 17;  //  CTRL key 
    public const long vbKeyMenu = 18;  //  MENU key 
    public const long vbKeyPause = 19;  //  PAUSE key 
    public const long vbKeyCapital = 20;  //  CAPS lock key
    public const long vbKeyEscape = 27;  //  ESC key 
    public const long vbKeySpace = 32;  //  SPACEBAR key 
    public const long vbKeyPageUp = 33;  //  PAGE UP key
    public const long vbKeyPageDown = 34;  //  PAGE DOWN key
    public const long vbKeyEnd = 35;  //  END key 
    public const long vbKeyHome = 36;  //  HOME key 
    public const long vbKeyLeft = 37;  //  LEFT ARROW key
    public const long vbKeyUp = 38;  //  UP ARROW key
    public const long vbKeyRight = 39;  //  RIGHT ARROW key
    public const long vbKeyDown = 40;  //  DOWN ARROW key
    public const long vbKeySelect = 41;  //  SELECT key 
    public const long vbKeyPrint = 42;  //  print SCREEN key
    public const long vbKeyExecute = 43;  //  EXECUTE key 
    public const long vbKeySnapshot = 44;  //  SNAPSHOT key 
    public const long vbKeyInsert = 45;  //  INS key 
    public const long vbKeyDelete = 46;  //  DEL key 
    public const long vbKeyHelp = 47;  // NUM HELP key 
    public const long vbKeyNumlock = 144;  //  lock key 
    public const long vbKeyA = 65;  //  A key 
    public const long vbKeyB = 66;  //  B key 
    public const long vbKeyC = 67;  //  C key 
    public const long vbKeyD = 68;  //  D key 
    public const long vbKeyE = 69;  //  E key 
    public const long vbKeyF = 70;  //  F key 
    public const long vbKeyG = 71;  //  G key 
    public const long vbKeyH = 72;  //  H key 
    public const long vbKeyI = 73;  //  I key 
    public const long vbKeyJ = 74;  //  J key 
    public const long vbKeyK = 75;  //  K key 
    public const long vbKeyL = 76;  //  L key 
    public const long vbKeyM = 77;  //  M key 
    public const long vbKeyN = 78;  //  N key 
    public const long vbKeyO = 79;  //  O key 
    public const long vbKeyP = 80;  //  P key 
    public const long vbKeyQ = 81;  //  Q key 
    public const long vbKeyR = 82;  //  R key 
    public const long vbKeyS = 83;  //  S key 
    public const long vbKeyT = 84;  //  T key 
    public const long vbKeyU = 85;  //  U key 
    public const long vbKeyV = 86;  //  V key 
    public const long vbKeyW = 87;  //  W key 
    public const long vbKeyX = 88;  //  X key 
    public const long vbKeyY = 89;  //  Y key 
    public const long vbKeyZ = 90;  //  Z key 
    public const long vbKey0 = 48;  //  0 key 
    public const long vbKey1 = 49;  //  1 key 
    public const long vbKey2 = 50;  //  2 key 
    public const long vbKey3 = 51;  //  3 key 
    public const long vbKey4 = 52;  //  4 key 
    public const long vbKey5 = 53;  //  5 key 
    public const long vbKey6 = 54;  //  6 key 
    public const long vbKey7 = 55;  //  7 key 
    public const long vbKey8 = 56;  //  8 key 
    public const long vbKey9 = 57;  //  9 key 
    public const long vbKeyNumpad0 = 96;  //  0 key 
    public const long vbKeyNumpad1 = 97;  //  1 key 
    public const long vbKeyNumpad2 = 98;  //  2 key 
    public const long vbKeyNumpad3 = 99;  // 4 3 key 
    public const long vbKeyNumpad4 = 100;  // 5 key  
    public const long vbKeyNumpad5 = 101;  // 6 key  
    public const long vbKeyNumpad6 = 102;  // 7 key  
    public const long vbKeyNumpad7 = 103;  // 8 key  
    public const long vbKeyNumpad8 = 104;  // 9 key  
    public const long vbKeyNumpad9 = 105;  // MULTIPLICATION key  
    public const long vbKeyMultiply = 106;  // PLUS SIGN (*) key
    public const long vbKeyAdd = 107;  // ENTER SIGN (+) key
    public const long vbKeySeparator = 108;  // MINUS (keypad) key 
    public const long vbKeySubtract = 109;  // DECIMAL SIGN (-) key
    public const long vbKeyDecimal = 110;  // DIVISION POINT(.) key 
    public const long vbKeyDivide = 111;  // F1 SIGN (/) key
    public const long vbKeyF1 = 112;  // F2 key  
    public const long vbKeyF2 = 113;  // F3 key  
    public const long vbKeyF3 = 114;  // F4 key  
    public const long vbKeyF4 = 115;  // F5 key  
    public const long vbKeyF5 = 116;  // F6 key  
    public const long vbKeyF6 = 117;  // F7 key  
    public const long vbKeyF7 = 118;  // F8 key  
    public const long vbKeyF8 = 119;  // F9 key  
    public const long vbKeyF9 = 120;  // F10 key  
    public const long vbKeyF10 = 121;  // F11 key  
    public const long vbKeyF11 = 122;  // F12 key  
    public const long vbKeyF12 = 123;  // F13 key  
    public const long vbKeyF13 = 124;  // F14 key  
    public const long vbKeyF14 = 125;  // F15 key  
    public const long vbKeyF15 = 126;  // F16 key  
    public const long vbKeyF16 = 127;  //  key  
    public const long vbBlack = 0x0;  // BLACK
    public const long vbBlue = 0x0000FF;  // BLUE
    public const long vbCyan = 0x00FFFF;  // CYAN
    public const long vbGreen = 0x00FF00;  // GREEN
    public const long vbMagenta = 0xFFFF00;  // MAGENTA
    public const long vbRed = 0xFF0000;  // RED
    public const long vbWhite = 0xFFFFFF;  // WHITE
    public const long vbYellow = 0xFF00FF;  // YELLOW
    public const long vbModal = 0x1;
    public enum AlignConstants {  vbAlignNone = 0, vbAlignTop = 1, vbAlignBottom = 2, vbAlignLeft = 3, vbAlignRight = 4 }
 }
