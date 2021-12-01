Attribute VB_Name = "modRegEx"
Option Explicit

Private mRegEx As Object
Private Property Get RegEx() As Object
  If mRegEx Is Nothing Then Set mRegEx = CreateObject("vbscript.regexp"): mRegEx.Global = True
  Set RegEx = mRegEx
End Property

Public Function RegExTest(ByVal Src As String, ByVal Find As String) As Boolean
On Error Resume Next
  RegEx.Pattern = Find
  RegExTest = RegEx.test(Src)
End Function

Public Function RegExCount(ByVal Src As String, ByVal Find As String) As Long
On Error Resume Next
  RegEx.Pattern = Find
  RegEx.Global = True
  RegExCount = RegEx.Execute(Src).Count
End Function

Public Function RegExNPos(ByVal Src As String, ByVal Find As String, Optional ByVal N As Long = 0) As Long
On Error Resume Next
  Dim RegM As Object, tempStr As String, tempStr2 As String
  RegEx.Pattern = Find
  RegEx.Global = True
  RegExNPos = RegEx.Execute(Src).Item(N).FirstIndex + 1
End Function

Public Function RegExNMatch(ByVal Src As String, ByVal Find As String, Optional ByVal N As Long = 0) As String
On Error Resume Next
  Dim RegM As Object, tempStr As String, tempStr2 As String
  RegEx.Pattern = Find
  RegEx.Global = True
  RegExNMatch = RegEx.Execute(Src).Item(N).Value
End Function

Public Function RegExReplace(ByVal Src As String, ByVal Find As String, ByVal Repl As String) As String
On Error Resume Next
  Dim RegM As Object, tempStr As String, tempStr2 As String
  RegEx.Pattern = Find
  RegEx.Global = True
  RegExReplace = RegEx.Replace(Src, Repl)
End Function

Public Function RegExSplit(ByVal szStr As String, ByVal szPattern As String) As Variant
On Error Resume Next
  Dim oAl As Variant, oRe As Variant, oMatches As Variant
  Set oRe = RegEx
  oRe.Pattern = "^(.*)(" & szPattern & ")(.*)$"
  oRe.IgnoreCase = True
  oRe.Global = True
  Set oAl = CreateObject("System.Collections.ArrayList")
  
  Do While True
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
  Dim T() As Variant
  T = RegExSplit(szStr, szPattern)
  RegExSplitCount = UBound(T) - LBound(T) + 1
End Function
