Attribute VB_Name = "modUtils"
Option Explicit

Public Function IsInStr(ByVal src As String, ByVal Find As String) As Boolean: IsInStr = InStr(src, Find) > 0: End Function
Public Function FileExists(ByVal Fn As String) As Boolean: FileExists = Dir(Fn) <> "": End Function
Public Function tLeft(ByVal Str As String, ByVal N As Long) As String: tLeft = Left(Trim(Str), N): End Function
Public Function tMid(ByVal Str As String, ByVal N As Long, Optional ByVal M As Long = 0) As String: tMid = IIf(M = 0, Mid(Trim(Str), N), Mid(Trim(Str), N, M)): End Function
Public Function StrCnt(ByVal src As String, ByVal Str As String) As Long: StrCnt = (Len(src) - Len(Replace(src, Str, ""))) / Len(Str): End Function
Public Function LMatch(ByVal src As String, ByVal tMatch As String) As Boolean: LMatch = Left(src, Len(tMatch)) = tMatch: End Function


Public Function sSpace(ByVal N As Long)
On Error Resume Next
  sSpace = Space(N)
End Function

Public Function nextBy(ByVal src As String, Optional ByVal Del As String = """", Optional ByVal Ind As Long = 1)
  Dim L As Long
  DoEvents
  L = InStr(src, Del)
  If L = 0 Then nextBy = IIf(Ind <= 1, src, ""): Exit Function
  Do While StrCnt(Left(src, L - 1), """") Mod 2 <> 0
    L = InStr(L + 1, src, Del)
    If L = 0 Then nextBy = IIf(Ind <= 1, src, ""): Exit Function
  Loop
  If Ind <= 1 Then
    nextBy = Left(src, L - 1)
  Else
    nextBy = nextBy(Mid(src, L + 1), Del, Ind - 1)
  End If
End Function

Public Function SplitWord(ByVal Source As String, Optional ByVal N As Long = 1, Optional ByVal Space As String = " ", Optional ByVal TrimResult As Boolean = True, Optional ByVal IncludeRest As Boolean = False) As String

'::::SplitWord
':::SUMMARY
': Return an indexed word from a string
':::DESCRIPTION
': Split()s a string based on a space (or other character) and return the word specified by the index.
': - Returns "" for 1 > N > Count
':::PARAMETERS
': - Source - The original source string to analyze
': - [N] = 1 - The index of the word to return (Default = 1)
': - [Space] = " " - The character to use as the "space" (defaults to %20).
': - [TrimResult] - Apply Trim() to the result (Default = True)
': - [IncludeRest] - Return the rest of the string starting at the indexed word (Default = False).
':::EXAMPLE
': - SplitWord("The Rain In Spain Falls Mostly", 4) == "Spain"
': - SplitWord("The Rain In Spain Falls Mostly", 4, , , True) == "Spain Falls Mostly"
': - SplitWord("a:b:c:d", -1, ":") === "d"
':::RETURN
':  String
':::SEE ALSO
': Split, CountWords
  Dim S, I As Long
  N = N - 1
  If Source = "" Then Exit Function
  S = Split(Source, Space)
  If N < 0 Then N = UBound(S) + N + 2
  If N < LBound(S) Or N > UBound(S) Then Exit Function
  If Not IncludeRest Then
    SplitWord = S(N)
  Else
    For I = N To UBound(S)
      SplitWord = SplitWord & IIf(Len(SplitWord) > 0, Space, "") & S(I)
    Next
  End If
  If TrimResult Then SplitWord = Trim(SplitWord)
End Function
Public Function CountWords(ByVal Source As String, Optional ByVal Space As String = " ") As Long
'::::CountWords
':::SUMMARY
': Returns the number of words in a string (determined by <Space> parameter)
':::DESCRIPTION
': Returns the count of words.
':::PARAMETERS
': - Source - The original source string to analyze
': - [Space] = " " - The character to use as the "space" (defaults to %20).
':::EXAMPLE
': - CountWords("The Rain In Spain Falls Mostly") == 6
': - CountWords("The Rain In Spain Falls Mostly", "n") == 4
':::RETURN
':  String
':::SEE ALSO
': SplitWord
  Dim L
' Count actual words.  Blank spaces don't count, before, after, or in the middle.
' Only a simple split and loop--there may be faster ways...
  For Each L In Split(Source, Space)
    If L <> "" Then CountWords = CountWords + 1
  Next
End Function

Public Function ArrSlice(ByRef sourceArray, ByVal fromIndex As Long, ByVal toIndex As Long)
  Dim Idx As Long
  Dim tempList()
  
  If Not IsArray(sourceArray) Then Exit Function
  
  fromIndex = FitRange(LBound(sourceArray), fromIndex, UBound(sourceArray))
  toIndex = FitRange(fromIndex, toIndex, UBound(sourceArray))
  
  For Idx = fromIndex To toIndex
    ArrAdd tempList, sourceArray(Idx)
  Next
  
  ArrSlice = tempList
End Function

Public Sub ArrAdd(ByRef Arr(), ByRef Item)
  Dim X As Long
  Err.Clear
On Error Resume Next
  X = UBound(Arr)
  If Err.Number <> 0 Then
    Arr = Array(Item)
    Exit Sub
  End If
  ReDim Preserve Arr(UBound(Arr) + 1)
  Arr(UBound(Arr)) = Item
End Sub
Public Function SubArr(ByVal sourceArray, ByVal fromIndex As Long, ByVal copyLength As Long)
  SubArr = ArrSlice(sourceArray, fromIndex, fromIndex + copyLength - 1)
End Function

Public Function InRange(ByVal LBnd, ByVal CHK, ByVal UBnd, Optional ByVal IncludeBounds As Boolean = True) As Boolean
On Error Resume Next  ' because we're doing this as variants..
  If IncludeBounds Then
    InRange = (CHK >= LBnd) And (CHK <= UBnd)
  Else
    InRange = (CHK > LBnd) And (CHK < UBnd)
  End If
End Function

Public Function FitRange(ByVal LBnd, ByVal CHK, ByVal UBnd)
On Error Resume Next
  If CHK < LBnd Then
    FitRange = LBnd
  ElseIf CHK > UBnd Then
    FitRange = UBnd
  Else
    FitRange = CHK
  End If
End Function

Public Function CodeSectionLoc(ByVal S As String) As Long
  Dim N As Long, K As Long

  Do
    N = InStr(S, "Attribute VB_Name")
    If N = 0 Then Exit Function
    K = InStr(N, S, vbLf)
    If K = 0 Then Exit Function
  Loop While Mid(S, K, 10) = "Attribute "
  
  CodeSectionLoc = K + 1
End Function

Public Function CodeSectionGlobalEndLoc(ByVal S As String)
  Dim A As Long, B As Long, C As Long
  A = InStr(S, "Function ")
  If A = 0 Then A = 1000000
  B = InStr(S, "Sub ")
  If B = 0 Then B = 1000000
  If B < A Then A = B
  B = InStr(S, "Property ")
  If B = 0 Then B = 1000000
  If B < A Then A = B
  
  If Mid(S, A - 8, 8) = "Private " Then A = A - 8
  If Mid(S, A - 7, 7) = "Public " Then A = A - 7

  CodeSectionGlobalEndLoc = A
End Function

Public Function DebugFolder() As String
    Dim oWSHShell As Object
    Set oWSHShell = CreateObject("WScript.Shell")
    DebugFolder = oWSHShell.SpecialFolders("Desktop") & "\test\"
    Set oWSHShell = Nothing
End Function

Public Function isOperator(ByVal S As String) As Boolean
  Select Case Trim(S)
    Case "+", "-", "/", "*", "&", "<>", "<", ">", "<=", ">=", "=", "Mod", "And", "Or", "Xor": isOperator = True
    Case Else: isOperator = False
  End Select
End Function
