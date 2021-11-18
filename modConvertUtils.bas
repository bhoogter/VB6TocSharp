Attribute VB_Name = "modConvertUtils"
Option Explicit

Private EOLComment As String
Private mStrings As Collection
Private nStringCnt As Long

Private Const DeStringToken_Base1 As String = "STRING_"
Private Const DeStringToken_Base2 As String = "TOKEN_"
Public Const DeStringToken_Base As String = DeStringToken_Base1 & DeStringToken_Base2


Public Function DeComment(ByVal Str As String, Optional ByVal Discard As Boolean = False) As String
  Dim A As Long
  Dim T As String, U As String
  Dim C As String
  DeComment = Str
  A = InStr(Str, "'")
  If A = 0 Then Exit Function
  Do While True
    T = Left(Str, A - 1)
    U = Replace(T, """", "")
    If (Len(T) - Len(U)) Mod 2 = 0 Then Exit Do
    A = InStr(A + 1, Str, "'")
    If A = 0 Then Exit Function
  Loop
  If Not Discard Then EOLComment = Mid(Str, A + 1)
  DeComment = RTrim(Left(Str, A - 1))
End Function

Public Function ReComment(ByVal Str As String, Optional ByVal KeepVBComments As Boolean = False) As String
  Dim C As String
  Dim Pr As String
  Pr = IIf(KeepVBComments, "'", "//")
  If EOLComment = "" Then ReComment = Str: Exit Function
  C = Pr & EOLComment
  EOLComment = ""
  If Not IsInStr(Str, vbCrLf) Then
    ReComment = Str & IIf(Len(Str) = 0, "", " ") & C
  Else
    ReComment = Replace(Str, vbCrLf, C & vbCrLf, , 1)         ' Always leave on end of first line...
  End If
  If Left(LTrim(ReComment), 2) = Pr Then ReComment = LTrim(ReComment)
End Function

Public Sub InitDeString()
  Set mStrings = New Collection
  nStringCnt = 0
End Sub

Private Function DeStringToken(ByVal N As Long) As String
  DeStringToken = DeStringToken_Base & Format(N, "00000")
End Function

Public Function DeString(ByVal S As String) As String
  Const Q As String = """"
  Dim Token As String
  Dim A As Long, B As Long, C As Long
  Dim K As String
  If mStrings Is Nothing Then InitDeString
  
'If IsInStr(S, """ArCheck.chkShowB") Then Stop
  
  A = InStr(S, Q)
  C = A
  If A > 0 Then
MidQuote:
    B = InStr(C + 1, S, Q)
    If B > 0 Then
      If Mid(S, B + 1, 1) = Q Then
        C = B + 1
        GoTo MidQuote
      End If
      nStringCnt = nStringCnt + 1
      Token = DeStringToken(nStringCnt)
      K = Mid(S, A, B - A + 1)
      mStrings.Add K, Token
      S = Left(S, A - 1) & Token & Mid(S, B + 1)
      DeString = DeString(S)
      Exit Function
    End If
  End If
  DeString = S
End Function

Public Function ReString(ByVal Str As String, Optional ByVal doConvertString As Boolean = False) As String
  Dim I As Long, T As String, V As String
  For I = 1 To nStringCnt
    T = DeStringToken(I)
    V = mStrings.Item(T)
    If V <> "" And doConvertString Then
      If Left(V, 1) = """" And Right(V, 1) = """" Then
        V = """" & InternalConvertString(Mid(V, 2, Len(V) - 2)) & """"
      End If
    End If
    Str = Replace(Str, T, V)
  Next
  ReString = Str
End Function


Private Function InternalConvertString(ByVal S As String) As String
  S = Replace(S, "\", "\\")
  S = Replace(S, """""", "\""")
  InternalConvertString = S
End Function

