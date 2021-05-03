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
using static modGit;
using static modDirStack;
using static modShell;
using static VB2CS.Forms.frm;
using static VB2CS.Forms.frmConfig;


static class modShell {
// Option Explicit
public const int SW_HIDE = 0;
public const int SW_SHOWNORMAL = 1;
public const int SW_SHOWMINIMIZED = 2;
public const int SW_SHOWMAXIMIZED = 3;
public const int SW_SHOW = 5;
public const int SW_SHOWDEFAULT = 10;
public const dynamic CREATE_NO_WINDOW = 0x8000000;
public const dynamic INFINITE = -1&;
private static int LastProcessID = 0;
private const dynamic DIRSEP = "\\";
public const dynamic NORMAL_PRIORITY_CLASS = 0x20&;
public enum enSW {
  enSW_HIDE = 0,
  enSW_NORMAL = 1,
  enSW_MAXIMIZE = 3,
  enSW_MINIMIZE = 6
}
public class STARTUPINFO {
 public int Cb = 0;
 public string lpReserved = "";
 public string lpDesktop = "";
 public string lpTitle = "";
 public int dwX = 0;
 public int dwY = 0;
 public int dwXSize = 0;
 public int dwYSize = 0;
 public int dwXCountChars = 0;
 public int dwYCountChars = 0;
 public int dwFillAttribute = 0;
 public int dwFlags = 0;
 public int wShowWindow = 0;
 public int cbReserved2 = 0;
 public int lpReserved2 = 0;
 public int hStdInput = 0;
 public int hStdOutput = 0;
 public int hStdError = 0;
}
public class PROCESS_INFORMATION {
 public int hProcess = 0;
 public int hThread = 0;
 public int dwProcessId = 0;
 public int dwThreadId = 0;
}
[DllImport("kernel32.dll")] private static extern void  Sleep(int dwMilliseconds);
[DllImport("kernel32.dll")] private static extern int CreateProcessA(int lpApplicationName, string lpCommandLine, int lpProcessAttributes, int lpThreadAttributes, int bInheritHandles, int dwCreationFlags, int lpEnvironment, int lpCurrentDirectory, ref STARTUPINFO lpStartupInfo, ref PROCESS_INFORMATION lpProcessInformation);
[DllImport("kernel32.dll")] private static extern int WaitForSingleObject(int hHandle, int dwMilliseconds);
[DllImport("kernel32.dll")] private static extern bool CloseHandle(ref int hObject);
[DllImport("user32.dll")] private static extern int GetDesktopWindow();
[DllImport("shell32.dll", EntryPoint = "ShellExecuteA")] private static extern int ShellExecute(int hwnd, string lpOperation, string lpFile, string lpParameters, string lpDirectory, int nShowCmd);


public static string RunCmdToOutput(string cmd, out string ErrStr, bool AsAdmin= false) {
  string RunCmdToOutput = "";
  // TODO (not supported): On Error GoTo RunError
  string A = "";
  string B = "";
  string C = "";

  int tLen = 0;
  int Iter = 0;

  A = TempFile();
  B = TempFile();
  if (!AsAdmin) {
    ShellAndWait("cmd /c " + cmd + " 1> " + A + " 2> " + B, enSW.enSW_HIDE);
  } else {
    C = TempFile("", "tmp_", ".bat");
    WriteFile(C, cmd + " 1> " + A + " 2> " + B, true);
    RunFileAsAdmin(C);
  }

  Iter = 0;
  const int MaxIter = 10;
  while(true) {
    tLen = FileLen(A);
    Sleep(800);
    if (Iter > MaxIter || FileLen(A) == tLen) {
      break;
    }
    Iter = Iter + 1;
  }
  RunCmdToOutput = ReadEntireFileAndDelete(A);
  if (Iter > MaxIter) {
    RunCmdToOutput = RunCmdToOutput + vbCrLf2 + "<<< OUTPUT TRUNCATED >>>";
  }
  ErrStr = ReadEntireFileAndDelete(B);
  DeleteFileIfExists(C);
  return RunCmdToOutput;


RunError:;
  RunCmdToOutput = "";
  ErrStr = "ShellOut.RunCmdToOutput: Command Execution Error - [" + Err().Number + "] " + Err().Description;
  return RunCmdToOutput;
}

/*
' to allow for Shell.
' This routine shells out to another application and waits for it to exit.
*/
public static void ShellAndWait(ref dynamic AppToRun, enSW SW= enSW.enSW_NORMAL) {
  PROCESS_INFORMATION NameOfProc = null;

  STARTUPINFO NameStart = null;

  int rc = 0;


  // TODO (not supported): On Error GoTo ErrorRoutineErr
  NameStart.Cb = Len(NameStart);
  if (SW == enSW.enSW_HIDE) {
    rc = CreateProcessA(0&, AppToRun, 0&, 0&, CLng(SW), CREATE_NO_WINDOW, 0&, 0&, NameStart, NameOfProc);
  } else {
    rc = CreateProcessA(0&, AppToRun, 0&, 0&, CLng(SW), NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc);
  }
  LastProcessID = NameOfProc.dwProcessId;
  rc = WaitForSingleObject(NameOfProc.hProcess, INFINITE);
  rc = CloseHandle(NameOfProc.hProcess);

ErrorRoutineResume:;
return;

ErrorRoutineErr:;
  MsgBox("AppShell.Form1.ShellAndWait: " + Err + Error);
  // TODO (not supported):     Resume Next

}

public static string TempFile(string UseFolder= "", string UsePrefix= "tmp_", string Extension= ".tmp", bool TestWrite= true) {
  string TempFile = "";
  string FN = "";
  string Res = "";

  if (UseFolder != "" && !DirExists(UseFolder)) {
    UseFolder = "";
  }
  if (UseFolder == "") {
    UseFolder = App.Path + DIRSEP;
  }
  if (Right(UseFolder, 1) != DIRSEP) {
    UseFolder = UseFolder + DIRSEP;
  }
  FN = Replace(UsePrefix + CDbl(DateTime.Now;) + "_" + App.ThreadID + "_" + Random(999999), ".", "_");
  while(FileExists(UseFolder + FN + ".tmp")) {
    FN = FN + Chr(Random(25) + Asc("a"));
  }
  TempFile = UseFolder + FN + Extension;

  if (TestWrite) {
    // TODO (not supported): On Error GoTo TestWriteFailed
    WriteFile(TempFile, "TEST", true, true);
    // TODO (not supported): On Error GoTo TestReadFailed
    Res = ReadFile(TempFile);
    if (Res != "TEST") {
      MsgBox("Test write to temp file " + TempFile + " failed." + vbCrLf + "Result (Len=" + Len(Res) + "):" + vbCrLf + Res, vbCritical);
    }
    // TODO (not supported): On Error GoTo TestClearFailed
    File.Delete(TempFile);();
  }
  return TempFile;


TestWriteFailed:;
  MsgBox("Failed to write temp file " + TempFile + "." + vbCrLf + Err().Description, vbCritical);
  return TempFile;

TestReadFailed:;
  MsgBox("Failed to read temp file " + TempFile + "." + vbCrLf + Err().Description, vbCritical);
  return TempFile;

TestClearFailed:;
  if (Err().Number == 53) {
    Err().Clear();
    // TODO (not supported):     Resume Next
  }

//BFH20160627
// Jerry wanted this commented out.  Absolutely horrible idea.
//  If IsDevelopment Then
  MsgBox("Failed to clear temp file " + TempFile + "." + vbCrLf + Err().Description, vbCritical);
//  End If
  return TempFile;

  return TempFile;
}

public static void RunShellExecuteAdmin(string App, int nHwnd= 0, int WindowState= SW_SHOWNORMAL) {
  if (nHwnd == 0) {
    nHwnd = GetDesktopWindow();
  }
  LastProcessID = ShellExecute(nHwnd, "runas", App, vbNullString, vbNullString, WindowState);
//  ShellExecute nHwnd, "runas", App, Command & " /admin", vbNullString, SW_SHOWNORMAL
}

public static bool RunFileAsAdmin(string App, int nHwnd= 0, int WindowState= SW_SHOWNORMAL) {
  bool RunFileAsAdmin = false;
//  If Not IsWinXP Then
  RunShellExecuteAdmin(App, nHwnd, WindowState);
//  Else
//    ShellOut App
//  End If
  RunFileAsAdmin = true;
  return RunFileAsAdmin;
}
}
