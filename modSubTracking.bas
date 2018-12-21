Attribute VB_Name = "modSubTracking"
Option Explicit

Public Type Variable
  Name As String
  asType As String
  asArray As String
  Param As Boolean
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
End Type
  
Private Lockout As Boolean

Private Vars() As Variable
Private Props() As Property


Public Property Get Analyze() As Boolean
  Analyze = Lockout
End Property

Public Sub SubBegin(Optional ByVal setLockout As Boolean = False)
  Dim nVars() As Variable
  Vars = nVars
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

Public Sub SubParamDecl(ByVal P As String, ByVal asType As String, ByVal asArray As String, ByVal isParam As Boolean)
  If Lockout Then Exit Sub

  Dim K As Variable
  K.Name = P
  K.Param = isParam
On Error Resume Next
  ReDim Preserve Vars(UBound(Vars) + 1)
  With Vars(UBound(Vars))
    .Name = P
    .asType = asType
    .Param = isParam
  End With
End Sub

Public Sub SubParamAssign(ByVal P As String)
  If Lockout Then Exit Sub
  
  Dim K As Long
  K = SubParamIndex(P)
  If K >= 0 Then
    With Vars(K)
      .Assigned = True
      If Not .Used Then .AssignedBeforeUsed = True
    End With
  End If
End Sub


Public Sub SubParamUsed(ByVal P As String)
  If Lockout Then Exit Sub

  Dim K As Long
  K = SubParamIndex(P)
  If K >= 0 Then Vars(K).Used = True
  If K >= 0 Then
    With Vars(K)
      .Used = True
      If Not .Used Then .UsedBeforeAssigned = True
    End With
  End If
End Sub

Public Sub SubParamUsedList(ByVal S As String)
  Dim Sp, L
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
  Dim Pro As String, asPublic As Boolean
  Dim asFunc As Boolean
  Dim GSL As String, pName As String, pArgs As String, pArgName As String, pType As String
  
  Pro = SplitWord(S, 1, vbCr)
  
  S = nlTrim(Replace(S, Pro, ""))
  If Right(S, 12) = "End Property" Then S = nlTrim(Left(S, Len(S) - 12))

  
  If LMatch(Pro, "Public ") Then Pro = Mid(Pro, 8): asPublic = True ' if one is public, both are...
  If LMatch(Pro, "Private ") Then Pro = Mid(Pro, 9)
  If LMatch(Pro, "Property ") Then Pro = Mid(Pro, 10)
  If LMatch(Pro, "Get ") Then Pro = Mid(Pro, 5): GSL = "get"
  If LMatch(Pro, "Let ") Then Pro = Mid(Pro, 5): GSL = "let"
  If LMatch(Pro, "Set ") Then Pro = Mid(Pro, 5): GSL = "set"
  pName = RegExNMatch(Pro, patToken)
  Pro = Mid(Pro, Len(pName) + 1)
  If LMatch(Pro, "(") Then Pro = Mid(Pro, 2)
  pArgs = nextBy(Pro, ")")
  If GSL = "get" And pArgs <> "" Or GSL <> "get" And InStr(pArgs, ",") > 0 Then asFunc = True
  If GSL = "set" Or GSL = "let" Then
    pArgName = SplitWord(pArgs, 1)
    If pArgName = "ByVal" Or pArgName = "ByRef" Then pArgName = SplitWord(pArgs, 2)
  End If
  Pro = Mid(Pro, Len(pArgs) + 1)
  If LMatch(Pro, ")") Then Pro = Trim(Mid(Pro, 2))
  If LMatch(Pro, "As ") Then
    Pro = Mid(Pro, 4)
    pType = Pro
  End If
  

  X = PropIndex(pName)
  If X = -1 Then
    X = 0
On Error Resume Next
    X = UBound(Props) + 1
On Error GoTo 0
  ReDim Preserve Props(X)
  End If
  
  PP = Props(X)

  PP.Name = pName
  If asPublic Then PP.asPublic = True  ' if one is public, both are...
  Select Case GSL
    Case "get"
                        PP.Getter = ConvertSub(S, , vbFalse)
                        PP.asType = ConvertDataType(pType)
    Case "set", "let":  PP.Setter = ConvertSub(S, , vbFalse)
                        PP.origArgName = pArgName
  End Select
End Sub

Public Function ReadOutProperties() As String
On Error Resume Next
  Dim I As Long, R As String, P As Property
  Dim T As String
  R = ""
  For I = LBound(Props) To UBound(Props)
    With Props(I)
      If .Name <> "" And Not (.Getter = "" And .Setter = "") Then
        If Not .asFunc Then
          If .asPublic Then R = R & "public "
          If .Getter = "" Then R = R & "writeonly "
          If .Setter = "" Then R = R & "readonly "
          R = R & .asType & " " & .Name & " {" & vbCrLf
          If .Getter <> "" Then
            R = R & "  get {" & vbCrLf
            R = R & "    " & .asType & " " & .Name & ";" & vbCrLf
            T = .Getter
            T = Replace(T, "Exit Property", "return " & .Name & ";")
            R = R & "    " & T & vbCrLf
            R = R & "  }" & vbCrLf
          End If
          If .Setter <> "" Then
            R = R & "  set {" & vbCrLf
            T = .Setter
            T = ReplaceToken(T, "value", "valueOrig")
            T = Replace(T, .origArgName, "value")
            T = Replace(T, "Exit Property", "return;") & vbCrLf
            R = R & "    " & T & vbCrLf
            R = R & "  }" & vbCrLf
          End If
          R = R & "}" & vbCrLf2
        Else
        End If
      End If
    End With
  Next
  
  ReadOutProperties = R
End Function
