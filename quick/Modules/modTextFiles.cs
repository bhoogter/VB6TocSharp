using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.IO;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Strings;
using static modUtils;
using static VBExtension;
using  System.Linq;


static class modTextFiles
{
    // @NO-LINT-DEPR
    // ::::modTextFiles
    // :::SUMMARY
    // : A processing module for text files.
    // :
    // :::DESCRIPTION
    // : Straight-forward, disposable methods for using text files.  Drastically reduces the complexity required to interact
    // : with flat text files, abstracting the developer.
    // :
    // :::INTERFACE
    // ::Public Interface
    // :- ReadFile
    // :- WriteFile
    // :- CountLines
    // :- VBFileCountLines
    // :- VBFileCountLines_Stat
    // :- ReadEntireFile
    // :- ReadEntireFileAndDelete
    // :- TailFile
    // :- HeadFile
    // :
    // :::SEE ALSO
    // :    - modXML, modCSV, modPath
    public static dynamic mFSO = null;
    private static dynamic FSO
    {
        get
        {
            dynamic _FSO;
            if (mFSO == null) mFSO = CreateObject("Scripting.FileSystemObject");
            _FSO = mFSO;
            return _FSO;
        }
    }

    public static bool DeleteFileIfExists(string sFIle, bool bNoAttributeClearing = false)
    {
        bool _DeleteFileIfExists = false;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (!FileExists(sFIle)) return _DeleteFileIfExists;
        if (!bNoAttributeClearing) File.SetAttributes(sFIle, 0);
        if (FileExists(sFIle)) File.Delete(sFIle);
        // DeleteFileIfExists = FileExists(sFile)
        _DeleteFileIfExists = true;
        return _DeleteFileIfExists;
    }
    public static string ReadEntireFile(string tFileName) => File.ReadAllText(tFileName);

    public static string ReadEntireFileAndDelete(string tFileName)
    {
        string _ReadEntireFileAndDelete = "";
        // ::::ReadEntireFileAndDelete
        // :::SUMMARY
        // :Read an entire file and safely delete it..
        // :::DESCRIPTION
        // :Reads the full contents of the file and then safely deletes it.
        // :
        // :If the file does not exist, no error is thrown, and an empty string is returned.
        // :::PARAMETERS
        // :- tFileName - The name of the file to read.
        // :::RETURN
        // :  String - The string contents of the file.
        // :::SEE ALSO
        // :  ReadEntireFile
        // TODO: (NOT SUPPORTED): On Error Resume Next
        _ReadEntireFileAndDelete = ReadEntireFile(tFileName);
        File.Delete(tFileName);
        return _ReadEntireFileAndDelete;
    }
    public static string ReadFile(string tFileName, int Startline = 1, int NumLines = 0) => File.ReadLines(tFileName).Skip(Startline).Take(NumLines).First();
    public static int CountFileLines(string SourceFile, bool IgnoreBlank = false, string IgnorePrefix = "")
    {
        int _CountFileLines = 0;
        // ::::CountFileLines
        // :::SUMMARY
        // :Returns the number of lines in a given file.
        // :::DESCRIPTION
        // :Retruns the number of lines in a file, based on the number of vbCr characters.
        // :
        // :- vbLf is completely ignored.
        // :- Blank lines can be optionally ignored
        // :- A prefix (such as # or ') can also be omitted from the count.
        // :
        // :If the file does not exist, no error is thrown, and an empty string is returned.
        // :::PARAMETERS
        // :- Source - The name of the file to read.
        // :- IgnoreBlank - Ignore blank lines in count.  Set to False to count all lines.  Default == TRUE
        // :- IgnorePrefix - Specify a string prefix to ignore in the count.  Popular options are the VB comment character (') and the utility file comment character (#).
        // :::RETURN
        // :  Long - The number of lines.
        // :::SEE ALSO
        // :  WriteFile, ReadFile, VBFileCountLines, CountLines
        _CountFileLines = CountLines(ReadEntireFile(SourceFile), IgnoreBlank, IgnorePrefix);
        return _CountFileLines;
    }
    public static int CountLines(string Source, bool IgnoreBlank = true, string IgnorePrefix = "'")
    {
        int _CountLines = 0;
        // ::::CountLines
        // :::SUMMARY
        // :Returns the number of lines in a given string (not a file).
        // :::DESCRIPTION
        // :Retruns the number of lines in a string, based on the number of vbCr characters.
        // :
        // :- vbLf is completely ignored.
        // :- Blank lines can be optionally ignored
        // :- A prefix (such as # or ') can also be omitted from the count.
        // :
        // :If the file does not exist, no error is thrown, and an empty string is returned.
        // :::PARAMETERS
        // :- Source - The string to count lines in.
        // :- IgnoreBlank - Ignore blank lines in count.  Set to False to count all lines.  Default == TRUE
        // :- IgnorePrefix - Specify a string prefix to ignore in the count.  Popular options are the VB comment character (') and the utility file comment character (#).
        // :::RETURN
        // :  Long - The number of lines.
        // :::SEE ALSO
        // :  WriteFile, ReadFile, VBFileCountLines, CountFileLines, LineByNumber
        dynamic L = null;
        Source = Replace(Source, vbLf, "");
        foreach (var iterL in new List<string>(Split(Source, vbCr)))
        {
            L = iterL;
            if (Trim(L) == "" && IgnoreBlank)
            {
                // Don't count...
            }
            else if (IgnorePrefix != "" && Left(LTrim(L), Len(IgnorePrefix)) == IgnorePrefix)
            {
                // Don't count...
            }
            else
            {
                _CountLines = _CountLines + 1;
            }
        }
        return _CountLines;
    }
    public static string LineByNumber(string Source, int Startline, int NumLines = 0, string NL = vbCrLf)
    {
        string _LineByNumber = "";
        // ::::LineByNumber
        // :::SUMMARY
        // :Returns the line(s) specified by the <StartLine> and <NumLines> parameters from a given <Source> string.
        // :::DESCRIPTION
        // :Similar to ReadFile, but for a string.
        // :
        // :If the file does not exist, no error is thrown, and an empty string is returned.
        // :
        // :- Reading before or end of multi-line string returns empty string.
        // :- Reading from center of lines beyond end of lines returns as many lines as possible.
        // :- Passing <NumLines> set to zero (0) returns remainder of lines (if any).
        // :::PARAMETERS
        // :- Source - The string to count lines in.
        // :- Startline - Ignore blank lines in count.  Set to False to count all lines.  Default == TRUE
        // :- NumLines - Specify a string prefix to ignore in the count.  Popular options are the VB comment character (') and the utility file comment character (#).
        // :- NL - The New Line charater(s) to use.  Default = vbCrLf
        // :::RETURN
        // :  String - The string at the specified location.
        // :::SEE ALSO
        // :  WriteFile, ReadFile, VBFileCountLines, CountFileLines, CountLines
        int A = 0;
        int B = 0;
        int I = 0;
        A = 0;
        if (Startline <= 0) Startline = 1;
        if (Startline == 1)
        {
            A = 1;
        }
        else
        {
            for (I = 1; I <= Startline - 1; I += 1)
            {
                A = InStr(A + 1, Source, NL);
                if (A == 0) return _LineByNumber;
            }
            A = A + Len(NL);
        }
        B = A;
        if (Left(Mid(Source, A), Len(NL)) != NL)
        {
            for (I = 1; I <= NumLines; I += 1)
            {
                B = InStr(B + 1, Source, NL);
                if (B == 0) { _LineByNumber = Mid(Source, A); return _LineByNumber; }
            }
        }
        _LineByNumber = Mid(Source, A, B - A);
        return _LineByNumber;
    }
    public static bool VBFileCountLines(string tFileName, out int Totl, out int Code, out int Blnk, out int Cmnt)
    {
        bool _VBFileCountLines = false;
        // ::::VBFileCountLines
        // :::SUMMARY
        // :Count lines in a VB6 file.
        // :::DESCRIPTION
        // :Count number of lines in a VB6 file.  Specifically tailored to account for the given parameters for VB6 code files.
        // :
        // :Returns the total line count, plus a breakdown of the following:
        // :- Code - Non-blank, non-comment-starting.
        // :- Blank - Count of blank lines.
        // :- Comment - Count of lines which are 100% comment (first character is ').
        // :
        // :If the file does not exist, no error is thrown, and an empty string is returned.
        // :::PARAMETERS
        // :- tFileName - The name of the file to read.
        // :- [Totl] - ByRef.  Returns total number of lines in file.
        // :- [Code] - ByRef.  Returns total number of code lines in file.
        // :- [Blnk] - ByRef.  Returns total number of blank lines in file.
        // :- [Cmnt] - ByRef.  Returns total number of comment lines in file.
        // :::RETURN
        // :  String - The string contents of the file.
        // :::SEE ALSO
        // :  ReadEntireFile, WriteFile, CountLines, VBFileCountLines_Stat
        string S = "";
        int N = 0;
        Totl = 0;
        Code = 0;
        Blnk = 0;
        Cmnt = 0;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (!FileExists(tFileName))
        {
            return _VBFileCountLines;
        }
        S = ReadEntireFile(tFileName);
        Totl = CountLines(S, false, "");
        Code = CountLines(S);
        N = CountLines(S,true, "");
        Cmnt = N - Code;
        Blnk = Totl - N;
        _VBFileCountLines = true;
        return _VBFileCountLines;
    }
    public static void VBFileCountLines_Stat(string tFileName)
    {
        // ::::VBFileCountLines_Stat
        // :::SUMMARY
        // :Print line count statistics for a file.
        // :::DESCRIPTION
        // :Raises a message box showing the file line count numbers.
        // :
        // :::PARAMETERS
        // :- tFileName - The name of the file to read.
        // :::SEE ALSO
        // :  ReadEntireFile, WriteFile, CountLines, VBFileCountLines
        int T = 0;
        int C = 0;
        int B = 0;
        int M = 0;
        if (VBFileCountLines(tFileName, out T, out C, out B, out M))
        {
            //MsgBox("File Line Stat: " + vbCrLf + " Totl: " + T + vbCrLf + "Code: " + C + vbCrLf + "Blnk: " + B + vbCrLf + "Cmnt: " + M, vbMsgBoxRtlReading);
        }
        else
        {
            //MsgBox("File Not Found: " + tFileName);
        }
    }
    public static bool WriteFile(string File, string Str, bool OverWrite = false, bool PreventNL = false)
    {
        bool _WriteFile = false;
        // ::::WriteFile
        // :::SUMMARY
        // :Write the given string to a file.
        // :::DESCRIPTION
        // :Writes a given text string to a file.
        // :
        // :Text may or may not contain new lines (multi-line write supported).
        // :
        // :A New-line is appended by default if not specified in thes tring.
        // :::PARAMETERS
        // :- File - The name of the file to read.
        // :- str - The text to write to the file.  Can be an empty string (blank line).
        // :- [OverWrite] - Default is to append.  Set to TRUE to delete file before write (overwrite contents).
        // :- [PreventNL] - By default, the end of the string is checked for a new line.  Use this to write to a file without a new-line.
        // :::RETURN
        // :  Boolean - Returns True.
        // :::SEE ALSO
        // :  ReadEntireFile, WriteFile, CountLines
        int FNo = 0;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        //FNo = FreeFile;
        if (OverWrite)
        {
            Kill(File);
            //VBOpenFile(File, "Output", FNo); // TODO: (NOT SUPPORTED) VB File Access Suppressed.  Convert manually: Open File For Output As #FNo
        }
        else
        {
            //VBOpenFile(File, "Append", FNo); // TODO: (NOT SUPPORTED) VB File Access Suppressed.  Convert manually: Open File For Append As #FNo
        }
        if (PreventNL || Right(Str, 2) == vbCrLf)
        {
            //VBWriteFile("Print #FNo, Str;"); // TODO: (NOT SUPPORTED) VB File Access Suppressed.  Convert manually: Print #FNo, Str;
        }
        else
        {
            //VBWriteFile("Print #FNo, Str"); // TODO: (NOT SUPPORTED) VB File Access Suppressed.  Convert manually: Print #FNo, Str
        }
        VBCloseFile(FNo); // TODO: (NOT SUPPORTED) VB File Access Suppressed.  Convert manually: Close #FNo
        _WriteFile = true;
        return _WriteFile;
    }

}
