Attribute VB_Name = "modSubTracking"
Option Explicit

Public Type Variable
  Name As String
  asType As String
  asArray As String
  Param As Boolean
  RetVal As Boolean
  Assigned As Boolean
  Used As Boolean
  AssignedBeforeUsed As Boolean
  UsedBeforeAssigned As Boolean
End Type

Public Type Property
  Name As String
  asPublic As Boolean
  asType As String
  asFunc As Boolean
  Getter As String
  Setter As String
  origArgName As String
  funcArgs As String
  origProto As String
End Type

Private Lockout As Boolean

Private Vars() As Variable
Private Props() As Property


Public Property Get Analyze() As Boolean
  Analyze = Lockout
End Property

Public Sub SubBegin(Optional ByVal setLockout As Boolean = False)
  If Not setLockout Then
    Dim nVars() As Variable
    Vars = nVars
  End If
  
  Lockout = Lockout
End Sub

Private Function SubParamIndex(ByVal P As String) As Long
  On Error GoTo NoEntries
  For SubParamIndex = LBound(Vars) To UBound(Vars)
    If Vars(SubParamIndex).Name = P Then Exit Function
  Next
NoEntries:
  SubParamIndex = -1
End Function

Public Function SubParam(ByVal P As String) As Variable
On Error Resume Next
  SubParam = Vars(SubParamIndex(P))
End Function

Public Sub SubParamDecl(ByVal P As String, ByVal asType As String, ByVal asArray As String, ByVal isParam As Boolean, ByVal isReturn As Boolean)
  If Lockout Then Exit Sub

  Dim K As Variable, N As Long
  K.Name = P
  K.Param = isParam
On Error Resume Next
  N = 0
  N = UBound(Vars) + 1
  ReDim Preserve Vars(N)
  Vars(N).Name = P
  Vars(N).asType = asType
  Vars(N).Param = isParam
  Vars(N).RetVal = isReturn
  Vars(N).asArray = asArray
End Sub

Public Sub SubParamAssign(ByVal P As String)
  If Lockout Then Exit Sub
  
  Dim K As Long
  K = SubParamIndex(P)
  If K >= 0 Then
    Vars(K).Assigned = True
    If Not Vars(K).Used Then Vars(K).AssignedBeforeUsed = True
  End If
End Sub


Public Sub SubParamUsed(ByVal P As String)
  If Lockout Then Exit Sub

  Dim K As Long
  K = SubParamIndex(P)
  If K >= 0 Then
    Vars(K).Used = True
    If Not Vars(K).Assigned Then Vars(K).UsedBeforeAssigned = True
  End If
End Sub

Public Sub SubParamUsedList(ByVal S As String)
  Dim Sp() As String, L As Variant
  If Lockout Then Exit Sub
  
  Sp = Split(S, ",")
  For Each L In Sp
    If L <> "" Then SubParamUsed L
  Next
End Sub



Public Sub ClearProperties()
  Dim nProps() As Property
  Props = nProps
End Sub

Private Function PropIndex(ByVal P As String) As Long
  On Error GoTo NoEntries
  For PropIndex = LBound(Props) To UBound(Props)
    If Props(PropIndex).Name = P Then Exit Function
  Next
NoEntries:
  PropIndex = -1
End Function

Public Sub AddProperty(ByVal S As String)
  Dim X As Long, PP As Property
  Dim Pro As String, origProto As String, asPublic As Boolean
  Dim asFunc As Boolean
  Dim GSL As String, pName As String, pArgs As String, pArgName As String, pType As String
  
  Pro = SplitWord(S, 1, vbCr)
  origProto = Pro
  
  S = nlTrim(Replace(S, Pro, ""))
  If Right(S, 12) = "End Property" Then S = nlTrim(Left(S, Len(S) - 12))

  
  If LMatch(Pro, "Public ") Then Pro = Mid(Pro, 8): asPublic = True ' if one is public, both are...
  If LMatch(Pro, "Private ") Then Pro = Mid(Pro, 9)
  If LMatch(Pro, "Friend ") Then Pro = Mid(Pro, 8)
  If LMatch(Pro, "Property ") Then Pro = Mid(Pro, 10)

  If LMatch(Pro, "Get ") Then Pro = Mid(Pro, 5): GSL = "get"
  If LMatch(Pro, "Let ") Then Pro = Mid(Pro, 5): GSL = "let"
  If LMatch(Pro, "Set ") Then Pro = Mid(Pro, 5): GSL = "set"
  pName = RegExNMatch(Pro, patToken)
  Pro = Mid(Pro, Len(pName) + 1)
  If LMatch(Pro, "(") Then Pro = Mid(Pro, 2)
  pArgs = nextBy(Pro, ")")
  If (GSL = "get" And pArgs <> "") Or (GSL <> "get" And InStr(pArgs, ",") > 0) Then
    asFunc = True
  End If
  If GSL = "set" Or GSL = "let" Then
    Dim fArg As String
    fArg = Trim(SplitWord(pArgs, -1, ","))
    If LMatch(fArg, "ByVal ") Then fArg = Mid(fArg, 7)
    If LMatch(fArg, "ByRef ") Then fArg = Mid(fArg, 7)
    pArgName = SplitWord(fArg, 1)
    If SplitWord(fArg, 2, " ") = "As" Then pType = SplitWord(fArg, 3, " ") Else pType = "Variant"
  End If
  Pro = Mid(Pro, Len(pArgs) + 1)
  If LMatch(Pro, ")") Then Pro = Trim(Mid(Pro, 2))
  If LMatch(Pro, "As ") Then
    Pro = Mid(Pro, 4)
    pType = Pro
  End If
  
  If pType = "" Then pType = "Variant"
  

  X = PropIndex(pName)
  If X = -1 Then
    X = 0
On Error Resume Next
    X = UBound(Props) + 1
On Error GoTo 0
    ReDim Preserve Props(X)
  End If
  
  Props(X).Name = pName
  Props(X).origProto = origProto
  If asPublic Then Props(X).asPublic = True  ' if one is public, both are...
  Select Case GSL
    Case "get"
      Props(X).Getter = ConvertSub(S, , vbFalse)
      Props(X).asType = ConvertDataType(pType)
      Props(X).asFunc = asFunc
      Props(X).funcArgs = pArgs
    Case "set", "let":
      Props(X).Setter = ConvertSub(S, , vbFalse)
      Props(X).origArgName = pArgName
      If pType <> "" Then Props(X).asType = ConvertDataType(pType)
      If asFunc Then Props(X).asFunc = True
      If pArgs <> "" Then Props(X).funcArgs = pArgs
  End Select
End Sub

Public Function ReadOutProperties(Optional ByVal asModule As Boolean = False) As String
On Error Resume Next
  Dim I As Long, R As String, P As Property
  Dim N As String, M As String
  Dim T As String
  R = ""
  M = ""
  N = vbCrLf
  I = -1
  For I = LBound(Props) To UBound(Props)
    If I = -1 Then GoTo NoItems
    If Props(I).Name <> "" And Not (Props(I).Getter = "" And Props(I).Setter = "") Then
      If Props(I).asPublic Then R = R & "public "
      If asModule Then R = R & "static "
'          If .Getter = "" Then R = R & "writeonly "
'          If .Setter = "" Then R = R & "readonly "
      If Props(I).asFunc Then
        R = R & " // TODO: Arguments not allowed on properties: " & Props(I).funcArgs & vbCrLf
        R = R & " //       " + Props(I).origProto & vbCrLf
      End If
      R = R & M & Props(I).asType & " " & Props(I).Name
      R = R & " {"
      
      If Props(I).Getter <> "" Then
        R = R & N & "  get {"
        R = R & N & "    " & Props(I).asType & " " & Props(I).Name & ";"
        T = Props(I).Getter
        T = Replace(T, "Exit(Property)", "return " & Props(I).Name & ";")
        R = R & N & "    " & T
        R = R & N & "  return " & Props(I).Name & ";"
        R = R & N & "  }"
      End If
      If Props(I).Setter <> "" Then
        R = R & N & "  set {"
        T = Props(I).Setter
        T = ReplaceToken(T, "value", "valueOrig")
        T = Replace(T, Props(I).origArgName, "value")
        T = Replace(T, "Exit Property", "return;")
        R = R & N & "    " & T
        R = R & N & "  }"
      End If
      R = R & N & "}"
      R = R & N
    End If
  Next
NoItems:

  ReadOutProperties = R
End Function
