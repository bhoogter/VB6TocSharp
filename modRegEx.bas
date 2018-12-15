Attribute VB_Name = "modRegEx"
Option Explicit

Public Function RegExTest(ByVal Src As String, ByVal Find As String) As Boolean
On Error Resume Next
  Dim RegEx As Object
  Set RegEx = CreateObject("vbscript.regexp")
  RegEx.Pattern = Find
  RegExTest = RegEx.Test(Src)
End Function

Public Function RegExCount(ByVal Src As String, ByVal Find As String) As Long
On Error Resume Next
  Dim RegEx As Object
  Set RegEx = CreateObject("vbscript.regexp")
  RegEx.Pattern = Find
  RegEx.Global = True
  RegExCount = RegEx.Execute(Src).Count
End Function

Public Function RegExNPos(ByVal Src As String, ByVal Find As String, Optional ByVal N As Long = 0) As Long
On Error Resume Next
  Dim RegEx As Object, RegM As Object, tempStr As String, tempStr2 As String
  Set RegEx = CreateObject("vbscript.regexp")
  RegEx.Pattern = Find
  RegEx.Global = True
  RegExNPos = RegEx.Execute(Src).Item(N).FirstIndex + 1
End Function

Public Function RegExNMatch(ByVal Src As String, ByVal Find As String, Optional ByVal N As Long = 0) As String
On Error Resume Next
  Dim RegEx As Object, RegM As Object, tempStr As String, tempStr2 As String
  Set RegEx = CreateObject("vbscript.regexp")
  RegEx.Pattern = Find
  RegEx.Global = True
  RegExNMatch = RegEx.Execute(Src).Item(N).Value
End Function

Public Function RegExSplit(ByVal szStr As String, ByVal szPattern As String)
On Error Resume Next
  Dim oAl, oRe, oMatches
  Set oRe = CreateObject("vbscript.regexp")
  oRe.Pattern = "^(.*)(" & szPattern & ")(.*)$"
  oRe.IgnoreCase = True
  oRe.Global = True
  Set oAl = CreateObject("System.Collections.ArrayList")
  
  Do
      Set oMatches = oRe.Execute(szStr)
      If oMatches.Count > 0 Then
          oAl.Add oMatches(0).SubMatches(2)
          szStr = oMatches(0).SubMatches(0)
      Else
          oAl.Add szStr
          Exit Do
      End If
  Loop
  oAl.Reverse
  RegExSplit = oAl.ToArray
End Function

Public Function RegExSplitCount(ByVal szStr As String, ByVal szPattern As String) As Long
On Error Resume Next
  Dim T()
  T = RegExSplit(szStr, szPattern)
  RegExSplitCount = UBound(T) - LBound(T) + 1
End Function
