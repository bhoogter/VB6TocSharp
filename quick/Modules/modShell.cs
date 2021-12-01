using System;
using System.Runtime.InteropServices;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modTextFiles;
using static modUtils;
using static VBExtension;



static class modShell
{
    public const int SW_HIDE = 0;
    public const int SW_SHOWNORMAL = 1;
    public const int SW_SHOWMINIMIZED = 2;
    public const int SW_SHOWMAXIMIZED = 3;
    public const int SW_SHOW = 5;
    public const int SW_SHOWDEFAULT = 10;
    public const int CREATE_NO_WINDOW = 0x8000000;
    public const int INFINITE = -1;
    public static int LastProcessID = 0;
    public const string DIRSEP = "\\";
    public const int NORMAL_PRIORITY_CLASS = 0x20;
    public enum enSW
    {
        enSW_HIDE = 0
    , enSW_NORMAL = 1
    , enSW_MAXIMIZE = 3
    , enSW_MINIMIZE = 6
    }
    public class STARTUPINFO
    {
        public int Cb;
        public string lpReserved;
        public string lpDesktop;
        public string lpTitle;
        public int dwX;
        public int dwY;
        public int dwXSize;
        public int dwYSize;
        public int dwXCountChars;
        public int dwYCountChars;
        public int dwFillAttribute;
        public int dwFlags;
        public int wShowWindow;
        public int cbReserved2;
        public int lpReserved2;
        public int hStdInput;
        public int hStdOutput;
        public int hStdError;
    }
    public class PROCESS_INFORMATION
    {
        public int hProcess;
        public int hThread;
        public int dwProcessId;
        public int dwThreadId;
    }
    [DllImport("kernel32")]
    private static extern void Sleep(int dwMilliseconds);
    [DllImport("kernel32")]
    private static extern int CreateProcessA(int lpApplicationName, string lpCommandLine, int lpProcessAttributes, int lpThreadAttributes, int bInheritHandles, int dwCreationFlags, int lpEnvironment, int lpCurrentDirectory, ref STARTUPINFO lpStartupInfo, ref PROCESS_INFORMATION lpProcessInformation);
    [DllImport("kernel32")]
    private static extern int WaitForSingleObject(int hHandle, int dwMilliseconds);
    [DllImport("kernel32")]
    private static extern bool CloseHandle(ref int hObject);
    [DllImport("USER32")]
    private static extern int GetDesktopWindow();
    [DllImport("shell32")]
    private static extern int ShellExecute(int hwnd, string lpOperation, string lpFile, string lpParameters, string lpDirectory, int nShowCmd);
    public static string RunCmdToOutput(string Cmd, out string ErrStr, bool AsAdmin = false)
    {
        string _RunCmdToOutput = "";
        ErrStr = "";
        // TODO: (NOT SUPPORTED): On Error GoTo RunError
        string A = "";
        string B = "";
        string C = "";
        int tLen = 0;
        int Iter = 0;
        A = TempFile();
        B = TempFile();
        if (!AsAdmin)
        {
            ShellAndWait("cmd /c " + Cmd + " 1> " + A + " 2> " + B, enSW.enSW_HIDE);
        }
        else
        {
            C = TempFile(".bat");
            WriteFile(C, Cmd + " 1> " + A + " 2> " + B, true);
            RunFileAsAdmin(C, 0, (int)enSW.enSW_HIDE);
        }
        Iter = 0;
        int MaxIter = 10;
        while (true)
        {
            tLen = (int)FileLen(A);
            Sleep(800);
            if (Iter > MaxIter || FileLen(A) == tLen) break;
            Iter = Iter + 1;
        }
        _RunCmdToOutput = ReadEntireFileAndDelete(A);
        if (Iter > MaxIter) _RunCmdToOutput = _RunCmdToOutput + vbCrLf2 + "<<< OUTPUT TRUNCATED >>>";
        ErrStr = ReadEntireFileAndDelete(B);
        DeleteFileIfExists(C);
        return _RunCmdToOutput;
    RunError:;
        _RunCmdToOutput = "";
        ErrStr = "ShellOut.RunCmdToOutput: Command Execution Error - [" + Err().Number + "] " + Err().Description;
        return _RunCmdToOutput;
    }
    // to allow for Shell.
    // This routine shells out to another application and waits for it to exit.
    public static void ShellAndWait(string AppToRun, enSW SW = enSW.enSW_NORMAL)
    {
        PROCESS_INFORMATION NameOfProc = null;
        STARTUPINFO NameStart = null;
        int RC = 0;
        // TODO: (NOT SUPPORTED): On Error GoTo ErrorRoutineErr
        NameStart.Cb = Len(NameStart);
        if (SW == enSW.enSW_HIDE)
        {
            RC = CreateProcessA(0, AppToRun, 0, 0, CInt(SW), CREATE_NO_WINDOW, 0, 0, ref NameStart, ref NameOfProc);
        }
        else
        {
            RC = CreateProcessA(0, AppToRun, 0, 0, CInt(SW), NORMAL_PRIORITY_CLASS, 0, 0, ref NameStart, ref NameOfProc);
        }
        LastProcessID = NameOfProc.dwProcessId;
        RC = WaitForSingleObject(NameOfProc.hProcess, INFINITE);
        RC = CloseHandle(ref NameOfProc.hProcess) ? 1 : 0;
    ErrorRoutineResume:;
        return;
    ErrorRoutineErr:;
        MsgBox("AppShell.Form1.ShellAndWait [" + Err().Number + "]: " + Err().Description);
        // TODO: (NOT SUPPORTED): Resume Next
    }
    public static string TempFile(string UseFolder = "", string UsePrefix = "tmp_", string Extension = ".tmp", bool TestWrite = true)
    {
        string _TempFile = "";
        string FN = "";
        string Res = "";
        if (UseFolder != "" && !DirExists(UseFolder)) UseFolder = "";
        if (UseFolder == "") UseFolder = AppContext.BaseDirectory + DIRSEP;
        if (Right(UseFolder, 1) != DIRSEP) UseFolder = UseFolder + DIRSEP;
        FN = Replace(UsePrefix + CDbl(DateTime.Now) + "_" + 0 + "_" + Random(999999), ".", "_");
        while (FileExists(UseFolder + FN + ".tmp"))
        {
            FN = FN + Chr(Random(25) + Asc("a"));
        }
        _TempFile = UseFolder + FN + Extension;
        if (TestWrite)
        {
            // TODO: (NOT SUPPORTED): On Error GoTo TestWriteFailed
            WriteFile(_TempFile, "TEST", true, true);
            // TODO: (NOT SUPPORTED): On Error GoTo TestReadFailed
            Res = ReadFile(_TempFile);
            if (Res != "TEST") MsgBox("Test write to temp file " + _TempFile + " failed." + vbCrLf + "Result (Len=" + Len(Res) + "):" + vbCrLf + Res, vbCritical);
            // TODO: (NOT SUPPORTED): On Error GoTo TestClearFailed
            Kill(_TempFile);
        }
        return _TempFile;
    TestWriteFailed:;
        MsgBox("Failed to write temp file " + _TempFile + "." + vbCrLf + Err().Description, vbCritical);
        return _TempFile;
    TestReadFailed:;
        MsgBox("Failed to read temp file " + _TempFile + "." + vbCrLf + Err().Description, vbCritical);
        return _TempFile;
    TestClearFailed:;
        if (Err().Number == 53)
        {
            // TODO: (NOT SUPPORTED): Err.Clear
            // TODO: (NOT SUPPORTED): Resume Next
        }
        // BFH20160627
        // Jerry wanted this commented out.  Absolutely horrible idea.
        // If IsDevelopment Then
        MsgBox("Failed to clear temp file " + _TempFile + "." + vbCrLf + Err().Description, vbCritical);
        // End If
        return _TempFile;
        return _TempFile;
    }
    public static void RunShellExecuteAdmin(string App, int nHwnd = 0, int WindowState = SW_SHOWNORMAL)
    {
        if (nHwnd == 0) nHwnd = GetDesktopWindow();
        LastProcessID = ShellExecute(nHwnd, "runas", App, vbNullString, vbNullString, WindowState);
        // ShellExecute nHwnd, __S1, App, Command & __S2, vbNullString, SW_SHOWNORMAL
    }
    public static bool RunFileAsAdmin(string App, int nHwnd = 0, int WindowState = SW_SHOWNORMAL)
    {
        bool _RunFileAsAdmin = false;
        // If Not IsWinXP Then
        RunShellExecuteAdmin(App, nHwnd, WindowState);
        // Else
        // ShellOut App
        // End If
        _RunFileAsAdmin = true;
        return _RunFileAsAdmin;
    }

}
