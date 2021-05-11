Attribute VB_Name = "modTextFiles"
Option Explicit
'@NO-LINT-DEPR
'::::modTextFiles
':::SUMMARY
': A processing module for text files.
':
':::DESCRIPTION
': Straight-forward, disposable methods for using text files.  Drastically reduces the complexity required to interact
': with flat text files, abstracting the developer.
':
':::INTERFACE
'::Public Interface
':- ReadFile
':- WriteFile
':- CountLines
':- VBFileCountLines
':- VBFileCountLines_Stat
':- ReadEntireFile
':- ReadEntireFileAndDelete
':- TailFile
':- HeadFile
':
':::SEE ALSO
':    - modXML, modCSV, modPath

Private mFSO As Object
Private Property Get FSO() As Object
  If mFSO Is Nothing Then Set mFSO = CreateObject("Scripting.FileSystemObject")
  Set FSO = mFSO
End Property

Public Function DeleteFileIfExists(ByVal sFIle As String, Optional ByVal bNoAttributeClearing As Boolean = False) As Boolean
On Error Resume Next
  If Not FileExists(sFIle) Then Exit Function
  If Not bNoAttributeClearing Then SetAttr sFIle, 0
  If FileExists(sFIle) Then Kill sFIle
'  DeleteFileIfExists = FileExists(sFile)
  DeleteFileIfExists = True
End Function



Public Function ReadEntireFile(ByVal tFileName As String) As String
'::::ReadEntireFile
':::SUMMARY
':Read an entire file.
':::DESCRIPTION
':Reads  the full contents of a file and returns the value as a string (without modification).
':::PARAMETERS
':- tFileName - The name of the file to read.
':::RETURN
':  String - The string contents of the file.
':::SEE ALSO
':  ReadFile, WriteFile, ReadEntireFileAndDelete

On Error Resume Next
  ReadEntireFile = FSO.OpenTextFile(tFileName, 1).ReadAll
  
  If FileLen(tFileName) / 10 <> Len(ReadEntireFile) / 10 Then
    MsgBox "ReadEntireFile was short: " & FileLen(tFileName) & " vs " & Len(ReadEntireFile)
  End If
'
'  Dim intFile As Long
'  intFile = FreeFile
'On Error Resume Next
'  Open tFileName For Input As #intFile
'  ReadEntireFile = Input$(LOF(intFile), #intFile)  '  LOF returns Length of File
'  Close #intFile
End Function


Public Function ReadEntireFileAndDelete(ByVal tFileName As String) As String
'::::ReadEntireFileAndDelete
':::SUMMARY
':Read an entire file and safely delete it..
':::DESCRIPTION
':Reads the full contents of the file and then safely deletes it.
':
':If the file does not exist, no error is thrown, and an empty string is returned.
':::PARAMETERS
':- tFileName - The name of the file to read.
':::RETURN
':  String - The string contents of the file.
':::SEE ALSO
':  ReadEntireFile

On Error Resume Next
  ReadEntireFileAndDelete = ReadEntireFile(tFileName)
  Kill tFileName
End Function

Public Function ReadFile(ByVal tFileName As String, Optional ByVal Startline As Long = 1, Optional ByVal NumLines As Long = 0) As String ', Optional ByRef WasEOF As Boolean = False)
'::::ReadFile
':::SUMMARY
':Random Access Read a given file based on line number.
':::DESCRIPTION
':Reads the specified lines from a given file.
':
':If the file does not exist, no error is thrown, and an empty string is returned.
':::PARAMETERS
':- tFileName - The name of the file to read.
':- StartLine - The line number to begin reading (the first line is 1).  If you try to read beyond the end of the file, an empty string is returned.
':- NumLines - If passed, attempts to read the specified number of lines.  Reading beyond the end of the file simply returns as many lines as possible.  Zero means read rest of file.  Default is zero.
':- WasEOF - If EOF checking is required, this ByRef parameter can be passed and checked later.  True if the file's EOF was reached.  False otherwise.
':::RETURN
':  String - The string contents of the file.
':::SEE ALSO
':  ReadEntireFile, WriteFile, CountLines, TailFile, HeadFile
  Dim FNum As Long, Line As String, LineNum As Long, Count As Long
  Static CacheFileName As String
  Static CacheFileDate As String
  Static CacheFileLoad() As String
  
  If tFileName = "" Or Not FileExists(tFileName) Then
'    WasEOF = True
    Exit Function
  End If
  
  If tFileName = CacheFileName Then
    If FileDateTime(tFileName) <> CacheFileDate Then CacheFileName = ""
  End If
  
  If tFileName <> CacheFileName Then
    CacheFileName = tFileName
    CacheFileDate = FileDateTime(tFileName)
    CacheFileLoad = Split(Replace(ReadEntireFile(tFileName), vbLf, ""), vbCr)
  End If
  
  If Startline = 1 And NumLines = 0 Then
    ReadFile = Join(CacheFileLoad, vbCrLf)
  Else
    ReadFile = Join(SubArr(CacheFileLoad, Startline - 1, NumLines), vbCrLf)
'    ReadFile = LineByNumber(CacheFileLoad, Startline, NumLines)
  End If
  
  Exit Function
  
'  If Startline < 1 Then Startline = 1
'  LineNum = 0
'  FNum = FreeFile
'  Open tFileName For Input As #FNum
'  Do While Not EOF(FNum)
'    LineNum = LineNum + 1
'    Line Input #FNum, Line
'    If LineNum >= Startline Then
'      ReadFile = ReadFile & IIf(Len(ReadFile) > 0, vbCrLf, "") & Line
'      Count = Count + 1
'    End If
'    If NumLines > 0 And Count >= NumLines Then GoTo Done
''    DoEvents
'  Loop
''  WasEOF = True
'Done:
'  Close #FNum
End Function

Public Function CountFileLines(ByVal SourceFile As String, Optional ByVal IgnoreBlank As Boolean = False, Optional ByVal IgnorePrefix As String = "") As Long
'::::CountFileLines
':::SUMMARY
':Returns the number of lines in a given file.
':::DESCRIPTION
':Retruns the number of lines in a file, based on the number of vbCr characters.
':
':- vbLf is completely ignored.
':- Blank lines can be optionally ignored
':- A prefix (such as # or ') can also be omitted from the count.
':
':If the file does not exist, no error is thrown, and an empty string is returned.
':::PARAMETERS
':- Source - The name of the file to read.
':- IgnoreBlank - Ignore blank lines in count.  Set to False to count all lines.  Default == TRUE
':- IgnorePrefix - Specify a string prefix to ignore in the count.  Popular options are the VB comment character (') and the utility file comment character (#).
':::RETURN
':  Long - The number of lines.
':::SEE ALSO
':  WriteFile, ReadFile, VBFileCountLines, CountLines
  CountFileLines = CountLines(ReadEntireFile(SourceFile), IgnoreBlank, IgnorePrefix)
End Function

Public Function CountLines(ByVal Source As String, Optional ByVal IgnoreBlank As Boolean = True, Optional ByVal IgnorePrefix As String = "'") As Long
'::::CountLines
':::SUMMARY
':Returns the number of lines in a given string (not a file).
':::DESCRIPTION
':Retruns the number of lines in a string, based on the number of vbCr characters.
':
':- vbLf is completely ignored.
':- Blank lines can be optionally ignored
':- A prefix (such as # or ') can also be omitted from the count.
':
':If the file does not exist, no error is thrown, and an empty string is returned.
':::PARAMETERS
':- Source - The string to count lines in.
':- IgnoreBlank - Ignore blank lines in count.  Set to False to count all lines.  Default == TRUE
':- IgnorePrefix - Specify a string prefix to ignore in the count.  Popular options are the VB comment character (') and the utility file comment character (#).
':::RETURN
':  Long - The number of lines.
':::SEE ALSO
':  WriteFile, ReadFile, VBFileCountLines, CountFileLines, LineByNumber
  Dim L As Variant
  Source = Replace(Source, vbLf, "")
  For Each L In Split(Source, vbCr)
    If Trim(L) = "" And IgnoreBlank Then
      ' Don't count...
    ElseIf IgnorePrefix <> "" And Left(LTrim(L), Len(IgnorePrefix)) = IgnorePrefix Then
      ' Don't count...
    Else
      CountLines = CountLines + 1
    End If
  Next
End Function

Public Function LineByNumber(ByVal Source As String, ByVal Startline As Long, Optional ByVal NumLines As Long = 0, Optional ByVal NL As String = vbCrLf) As String
'::::LineByNumber
':::SUMMARY
':Returns the line(s) specified by the <StartLine> and <NumLines> parameters from a given <Source> string.
':::DESCRIPTION
':Similar to ReadFile, but for a string.
':
':If the file does not exist, no error is thrown, and an empty string is returned.
':
':- Reading before or end of multi-line string returns empty string.
':- Reading from center of lines beyond end of lines returns as many lines as possible.
':- Passing <NumLines> set to zero (0) returns remainder of lines (if any).
':::PARAMETERS
':- Source - The string to count lines in.
':- Startline - Ignore blank lines in count.  Set to False to count all lines.  Default == TRUE
':- NumLines - Specify a string prefix to ignore in the count.  Popular options are the VB comment character (') and the utility file comment character (#).
':- NL - The New Line charater(s) to use.  Default = vbCrLf
':::RETURN
':  String - The string at the specified location.
':::SEE ALSO
':  WriteFile, ReadFile, VBFileCountLines, CountFileLines, CountLines
  Dim A As Long, B As Long
  Dim I As Long
  A = 0
  If Startline <= 0 Then Startline = 1
  
  If Startline = 1 Then
    A = 1
  Else
    For I = 1 To Startline - 1
      A = InStr(A + 1, Source, NL)
      If A = 0 Then Exit Function
    Next
    A = A + Len(NL)
  End If
  
  B = A
  If Left(Mid(Source, A), Len(NL)) <> NL Then
    For I = 1 To NumLines
      B = InStr(B + 1, Source, NL)
      If B = 0 Then LineByNumber = Mid(Source, A): Exit Function
    Next
  End If
  
  LineByNumber = Mid(Source, A, B - A)
End Function



Public Function VBFileCountLines(ByVal tFileName As String, Optional ByRef Totl As Long, Optional ByRef Code As Long, Optional ByRef Blnk As Long, Optional ByRef Cmnt As Long) As Boolean
'::::VBFileCountLines
':::SUMMARY
':Count lines in a VB6 file.
':::DESCRIPTION
':Count number of lines in a VB6 file.  Specifically tailored to account for the given parameters for VB6 code files.
':
':Returns the total line count, plus a breakdown of the following:
':- Code - Non-blank, non-comment-starting.
':- Blank - Count of blank lines.
':- Comment - Count of lines which are 100% comment (first character is ').
':
':If the file does not exist, no error is thrown, and an empty string is returned.
':::PARAMETERS
':- tFileName - The name of the file to read.
':- [Totl] - ByRef.  Returns total number of lines in file.
':- [Code] - ByRef.  Returns total number of code lines in file.
':- [Blnk] - ByRef.  Returns total number of blank lines in file.
':- [Cmnt] - ByRef.  Returns total number of comment lines in file.
':::RETURN
':  String - The string contents of the file.
':::SEE ALSO
':  ReadEntireFile, WriteFile, CountLines, VBFileCountLines_Stat
  Dim S As String, N As Long
  Totl = 0
  Code = 0
  Blnk = 0
  Cmnt = 0
  
On Error Resume Next
  If Not FileExists(tFileName) Then
    Exit Function
  End If
  S = ReadEntireFile(tFileName)
  Totl = CountLines(S, False, "")
  Code = CountLines(S)
  N = CountLines(S, , "")
  Cmnt = N - Code
  Blnk = Totl - N
  VBFileCountLines = True
End Function

Public Sub VBFileCountLines_Stat(ByVal tFileName As String)
'::::VBFileCountLines_Stat
':::SUMMARY
':Print line count statistics for a file.
':::DESCRIPTION
':Raises a message box showing the file line count numbers.
':
':::PARAMETERS
':- tFileName - The name of the file to read.
':::SEE ALSO
':  ReadEntireFile, WriteFile, CountLines, VBFileCountLines
  Dim T As Long, C As Long, B As Long, M As Long
  If VBFileCountLines(tFileName, T, C, B, M) Then
    MsgBox "File Line Stat: " & vbCrLf & " Totl: " & T & vbCrLf & "Code: " & C & vbCrLf & "Blnk: " & B & vbCrLf & "Cmnt: " & M, vbMsgBoxRtlReading
  Else
    MsgBox "File Not Found: " & tFileName
  End If
End Sub


Public Function WriteFile(ByVal File As String, ByVal Str As String, Optional ByVal OverWrite As Boolean = False, Optional ByVal PreventNL As Boolean = False) As Boolean
'::::WriteFile
':::SUMMARY
':Write the given string to a file.
':::DESCRIPTION
':Writes a given text string to a file.
':
':Text may or may not contain new lines (multi-line write supported).
':
':A New-line is appended by default if not specified in thes tring.
':::PARAMETERS
':- File - The name of the file to read.
':- str - The text to write to the file.  Can be an empty string (blank line).
':- [OverWrite] - Default is to append.  Set to TRUE to delete file before write (overwrite contents).
':- [PreventNL] - By default, the end of the string is checked for a new line.  Use this to write to a file without a new-line.
':::RETURN
':  Boolean - Returns True.
':::SEE ALSO
':  ReadEntireFile, WriteFile, CountLines
  Dim FNo As Long
On Error Resume Next
  FNo = FreeFile
  If OverWrite Then
    Kill File
    Open File For Output As #FNo
  Else
    Open File For Append As #FNo
  End If
  If PreventNL Or Right(Str, 2) = vbCrLf Then
    Print #FNo, Str;
  Else
    Print #FNo, Str
  End If
  Close #FNo
  WriteFile = True
End Function

