Attribute VB_Name = "modUtils"
Option Explicit


Public Const patToken As String = "([a-zA-Z_][a-zA-Z_0-9]*)"
Public Const patNotToken As String = "([^a-zA-Z_0-9])"
Public Const patTokenDot As String = "([a-zA-Z_.][a-zA-Z_0-9.]*)"
Public Const vbCrLf2 As String = vbCrLf & vbCrLf
Public Const vbCrLf3 As String = vbCrLf & vbCrLf & vbCrLf
Public Const vbCrLf4 As String = vbCrLf & vbCrLf & vbCrLf & vbCrLf


Public Function IsInStr(ByVal Src As String, ByVal Find As String) As Boolean: IsInStr = InStr(Src, Find) > 0: End Function
Public Function FileExists(ByVal Fn As String) As Boolean: FileExists = Dir(Fn) <> "": End Function
Public Function FileName(ByVal Fn As String) As String: FileName = Mid(Fn, InStrRev(Fn, "\") + 1): End Function
Public Function FilePath(ByVal Fn As String) As String: FilePath = Left(Fn, InStrRev(Fn, "\")): End Function
Public Function ChgExt(ByVal Fn As String, ByVal NewExt As String) As String: ChgExt = Left(Fn, InStrRev(Fn, ".") - 1) & NewExt: End Function
Public Function tLeft(ByVal Str As String, ByVal N As Long) As String: tLeft = Left(Trim(Str), N): End Function
Public Function tMid(ByVal Str As String, ByVal N As Long, Optional ByVal M As Long = 0) As String: tMid = IIf(M = 0, Mid(Trim(Str), N), Mid(Trim(Str), N, M)): End Function
Public Function StrCnt(ByVal Src As String, ByVal Str As String) As Long: StrCnt = (Len(Src) - Len(Replace(Src, Str, ""))) / Len(Str): End Function
Public Function LMatch(ByVal Src As String, ByVal tMatch As String) As Boolean: LMatch = Left(Src, Len(tMatch)) = tMatch: End Function
Public Function tLMatch(ByVal Src As String, ByVal tMatch As String) As Boolean: tLMatch = Left(LTrim(Src), Len(tMatch)) = tMatch: End Function
Public Function Px(ByVal Twips As Long) As Long:  Px = Twips / 14: End Function
Public Function Quote(ByVal S As String) As String:  Quote = """" & S & """": End Function

Public Function WriteOut(ByVal F As String, ByVal S As String, Optional ByVal O As String = "") As Boolean: WriteOut = WriteFile(OutputFolder(O) & F, S, True): End Function

Public Function FileExt(ByVal Fn As String, Optional ByVal vLCase As Boolean = True) As String
  If Fn = "" Then Exit Function
  FileExt = Mid(Fn, InStrRev(Fn, "."))
  FileExt = IIf(vLCase, LCase(FileExt), FileExt)
End Function


Public Function deQuote(ByVal Src As String) As String
  If Left(Src, 1) = """" Then Src = Mid(Src, 2)
  If Right(Src, 1) = """" Then Src = Left(Src, Len(Src) - 1)
  deQuote = Src
End Function

Public Function deWS(ByVal S As String) As String
  Do While IsInStr(S, " " & vbCrLf)
    S = Replace(S, " " & vbCrLf, vbCrLf)
  Loop
  Do While IsInStr(S, vbCrLf4)
    S = Replace(S, vbCrLf4, vbCrLf3)
  Loop
  
  S = Replace(S, "{" & vbCrLf2, "{" & vbCrLf)
  S = RegExReplace(S, "(" & vbCrLf2 & ")([ ]*{)", vbCrLf & "$2")
  deWS = S
End Function

Public Function nlTrim(ByVal Str As String)
  Do While InStr(" " & vbTab & vbCr & vbLf, Left(Str, 1)) <> 0 And Str <> "": Str = Mid(Str, 2): Loop
  Do While InStr(" " & vbTab & vbCr & vbLf, Right(Str, 1)) <> 0 And Str <> "": Str = Mid(Str, 1, Len(Str) - 1): Loop
  nlTrim = Str
End Function


Public Function sSpace(ByVal N As Long)
On Error Resume Next
  sSpace = Space(N)
End Function

Public Function nextBy(ByVal Src As String, Optional ByVal Del As String = """", Optional ByVal Ind As Long = 1, Optional ByVal ProcessVBComments As Boolean = False)
  Dim L As Long
  DoEvents
  L = InStr(Src, Del)
  If L = 0 Then nextBy = IIf(Ind <= 1, Src, ""): Exit Function
  Do While StrCnt(Left(Src, L - 1), """") Mod 2 <> 0
    L = InStr(L + 1, Src, Del)
    If L = 0 Then nextBy = IIf(Ind <= 1, Src, ""): Exit Function
  Loop
  If Ind <= 1 Then
    nextBy = Left(Src, L - 1)
  Else
    nextBy = nextBy(Mid(Src, L + 1), Del, Ind - 1)
  End If
End Function

Public Function ReplaceToken(ByVal Src As String, ByVal OrigToken As String, ByVal NewToken As String) As String
  ReplaceToken = RegExReplace(Src, "([^a-zA-Z_0-9])(" & OrigToken & ")([^a-zA-Z_0-9])", "$1" & NewToken & "$3")
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
  Const Token As String = "Attribute VB_Name"
  Dim N As Long, K As Long

  N = InStr(S, Token)
  If N = 0 Then Exit Function
  Do
    N = InStr(N, S, vbLf) + 1
    If N <= 1 Then Exit Function
  Loop While Mid(S, N, 10) = "Attribute "
  
  CodeSectionLoc = N
End Function

Public Function CodeSectionGlobalEndLoc(ByVal S As String)
  Do
    CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc + RegExNPos(Mid(S, CodeSectionGlobalEndLoc + 1), "([^a-zA-Z0-9_]Function |[^a-zA-Z0-9_]Sub |[^a-zA-Z0-9_]Property )") + 1
    If CodeSectionGlobalEndLoc = 1 Then CodeSectionGlobalEndLoc = Len(S): Exit Function
  Loop While Mid(S, CodeSectionGlobalEndLoc - 8, 8) = "Declare "
  If CodeSectionGlobalEndLoc >= 8 Then
    If Mid(S, CodeSectionGlobalEndLoc - 7, 7) = "Public " Then CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc - 7
    If Mid(S, CodeSectionGlobalEndLoc - 8, 8) = "Private " Then CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc - 8
  End If
  CodeSectionGlobalEndLoc = CodeSectionGlobalEndLoc - 1
End Function

Public Function OutputSubFolder(ByVal F As String) As String
  Select Case FileExt(F)
    Case ".bas": OutputSubFolder = "Modules\"
    Case ".cls": OutputSubFolder = "Classes\"
    Case ".frm": OutputSubFolder = "Forms\"
    Case Else:   OutputSubFolder = ""
  End Select
End Function

Public Function OutputFolder(Optional ByVal F As String) As String
  OutputFolder = "C:\Users\benja\workspace\VS2017\WinCDS.NET\WinCDS.NET\" & OutputSubFolder(F)
End Function



Public Function isOperator(ByVal S As String) As Boolean
  Select Case Trim(S)
    Case "+", "-", "/", "*", "&", "<>", "<", ">", "<=", ">=", "=", "Mod", "And", "Or", "Xor": isOperator = True
    Case Else: isOperator = False
  End Select
End Function

Public Function Prg(Optional ByVal Val As Long = -1, Optional ByVal Max As Long = -1, Optional ByVal Cap = "#")
  frm.Prg Val, Max, Cap
End Function

Public Function cVal(coll As Collection, Key As String, Optional ByVal Def As String = "") As String
  On Error Resume Next
  cVal = Def
  cVal = coll.Item(LCase(Key))
End Function

Public Function cValP(coll As Collection, Key As String, Optional ByVal Def As String = "") As String
  cValP = P(deQuote(cVal(coll, Key, Def)))
End Function

Public Function P(ByVal Str As String) As String
  Str = Replace(Str, "&", "&amp;")
  Str = Replace(Str, "<", "&lt;")
  Str = Replace(Str, ">", "&gt;")
  P = Str
End Function

Public Function ModuleName(ByVal S As String) As String
  Dim J As Long, K As Long
  Const NameTag As String = "Attribute VB_Name = """
  J = InStr(S, NameTag) + Len(NameTag)
  K = InStr(J, S, """") - J
  ModuleName = Mid(S, J, K)
End Function

Public Function IsInCode(ByVal Src As String, ByVal N As Long)
  Dim I As Long, C As String
  Dim Qu As Boolean
  IsInCode = False
  For I = N To 1 Step -1
    C = Mid(Src, I, 1)
    If C = vbCr Or C = vbLf Then
      IsInCode = True
      Exit Function
    ElseIf C = """" Then
      Qu = Not Qu
    ElseIf C = "'" Then
      If Not Qu Then Exit Function
    End If
  Next
  IsInCode = True
End Function

Public Function TokenList(ByVal S As String) As String
  Dim I As Long, N As Long, T As String
  N = RegExCount(S, patToken)
  For I = 0 To N - 1
    T = RegExNMatch(S, patToken, I)
    TokenList = TokenList & "," & T
  Next
End Function

Public Function Random(Optional ByVal Max As Long = 10000) As Long
  Randomize
  Random = ((Rnd * Max) + 1)
End Function
