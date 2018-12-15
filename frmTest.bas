Attribute VB_Name = "frmTest"
Option Explicit


Public Function TestCase() As String
  Dim A As String, N As String, M As String
  A = "": M = "": N = vbCrLf
  
  A = A & M & "Public Function ReadEntireFile(ByVal FileName As String) As String"
  A = A & N & "'::::ReadEntireFile"
  A = A & N & "':::SUMMARY"
  A = A & N & "':Read an entire file."
  A = A & N & "':::DESCRIPTION"
  A = A & N & "':Reads  the full contents of a file and returns the value as a string (without modification)."
  A = A & N & "':::PARAMETERS"
  A = A & N & "':- FileName - The name of the file to read."
  A = A & N & "':::RETURN"
  A = A & N & "':  String - The string contents of the file."
  A = A & N & "':::SEE ALSO"
  A = A & N & "':  ReadFile, WriteFile, ReadEntireFileAndDelete"
  A = A & N & ""
  A = A & N & "On Error Resume Next"
  A = A & N & "  With CreateObject(""Scripting.FileSystemObject"")"
  A = A & N & "    ReadEntireFile = .OpenTextFile(FileName, 1).ReadAll"
  A = A & N & "  End With"
  A = A & N & "  "
  A = A & N & "  If FileLen(FileName) / 10 <> Len(ReadEntireFile) / 10 Then"
  A = A & N & "    MsgBox ""ReadEntireFile was short: "" & FileLen(FileName) & "" vs "" & Len(ReadEntireFile)"
  A = A & N & "  End If"
  A = A & N & "'"
  A = A & N & "'  Dim intFile As Long"
  A = A & N & "'  intFile = FreeFile"
  A = A & N & "'On Error Resume Next"
  A = A & N & "'  Open FileName For Input As #intFile"
  A = A & N & "'  ReadEntireFile = Input$(LOF(intFile), #intFile)  '  LOF returns Length of File"
  A = A & N & "'  Close #intFile"
  A = A & N & "End Function"
  TestCase = A
End Function

Public Function TestCase2()
  Dim A As String, N As String, M As String
  A = "": M = "": N = vbCrLf
  
A = A & N & "Public Function DescribeColor(ByVal RGB As Long) As String"
A = A & N & "  Dim R As Long, G As Long, B As Long"
A = A & N & "  Select Case RGB"
A = A & N & "    Case vbBlack:     DescribeColor = ""BLACK"""
A = A & N & "    Case vbBlue:      DescribeColor = ""BLUE"""
A = A & N & "    Case vbCyan:      DescribeColor = ""CYAN"""
A = A & N & "    Case vbGreen:     DescribeColor = ""GREEN"""
A = A & N & "    Case vbMagenta:   DescribeColor = ""MAGENTA"""
A = A & N & "    Case vbRed:       DescribeColor = ""RED"""
A = A & N & "    Case vbWhite:     DescribeColor = ""WHITE"""
A = A & N & "    Case vbYellow:    DescribeColor = ""YELLOW"""
A = A & N & ""
A = A & N & "    Case Else"
A = A & N & "      DescribeColor = ""OTHER"""
A = A & N & "      R = RGB And 255"
A = A & N & "      G = (RGB And 65280) / 256"
A = A & N & "      B = (RGB And 16711680) / 65536"
A = A & N & "      DescribeColor = DescribeColor & "" (R:"" & R & "",G:"" & G & "",:"" & B & "")"""
A = A & N & "  End Select"
A = A & N & ""
A = A & N & " ' If SingleLineIf Then DoSomeStuff: AndMore()"
A = A & N & ""
A = A & N & "GotoLabel:"
A = A & N & "  SomeFunction()"
A = A & N & "End Function"
  
  TestCase2 = SanitizeCode(A)
End Function
