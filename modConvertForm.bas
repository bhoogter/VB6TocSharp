Attribute VB_Name = "modConvertForm"
Option Explicit

Private EventStubs As String

Public Function Frm2Xml(ByVal F As String) As String
  Dim Sp, L, I As Long
  Dim R As String
  Sp = Split(F, vbCrLf)
  
  For Each L In Sp
    L = Trim(L)
    If L = "" Then GoTo NextLine
    If Left(L, 10) = "Attribute " Or Left(L, 8) = "VERSION " Then
    ElseIf Left(L, 6) = "Begin " Then
      R = R & sSpace(I * SpIndent) & "<item type=""" & SplitWord(L, 2) & """ name=""" & SplitWord(L, 3) & """>" & vbCrLf
      I = I + 1
    ElseIf L = "End" Then
      I = I - 1
      R = R & sSpace(I * SpIndent) & "</item>" & vbCrLf
    Else
      R = R & sSpace(I * SpIndent) & "<prop name=""" & SplitWord(L, 1, "=") & """ value=""" & SplitWord(L, 2, "=", , True) & """ />" & vbCrLf
    End If
NextLine:
  Next
  Frm2Xml = R
End Function

Public Function FormControls(ByVal Src As String, ByVal F As String, Optional ByVal asLocal As Boolean = True)
  Dim Sp, L, I As Long
  Dim R As String, T As String
  Dim Nm As String, Ty As String
  Sp = Split(F, vbCrLf)
  
  For Each L In Sp
    L = Trim(L)
    If L = "" Then GoTo NextLine
    If Left(L, 6) = "Begin " Then
      Ty = SplitWord(L, 2)
      Nm = SplitWord(L, 3)
      Select Case Ty
        Case "VB.Form"
        Case Else
          T = Src & ":" & IIf(asLocal, "", Src & ".") & Nm & ":Control:" & Ty
          If Right(R, Len(T)) <> T Then R = R & vbCrLf & T
      End Select
    End If
NextLine:
  Next
  FormControls = R
End Function

Public Function ConvertFormUi(ByVal F As String, ByVal CodeSection As String) As String
  Dim Stck(0 To 100)
  Dim Sp, L, J As Long, K As Long, I As Long, Tag As String
  Dim M As String
  Dim R As String
  Dim Prefix As String
  Dim Props As Collection, pK As String, pV As String
  Sp = Split(F, vbCrLf)
  
  EventStubs = ""
  
  For K = LBound(Sp) To UBound(Sp)
    L = Trim(Sp(K))
    If L = "" Then GoTo NextLine
    
    If Left(L, 10) = "Attribute " Or Left(L, 8) = "VERSION " Then
    ElseIf Left(L, 6) = "Begin " Then
      Set Props = New Collection
      J = 0
      Do
        J = J + 1
        M = Trim(Sp(K + J))
        If LMatch(M, "Begin ") Or M = "End" Then Exit Do
        
        If LMatch(M, "BeginProperty ") Then
          Prefix = LCase(Prefix & SplitWord(M, 2) & ".")
        ElseIf LMatch(M, "EndProperty") Then
          Prefix = Left(Prefix, Len(Prefix) - 1)
          If Not IsInStr(Prefix, ".") Then
            Prefix = ""
          Else
            Prefix = Left(Prefix, InStrRev(Left(Prefix, Len(Prefix) - 1), "."))
          End If
        Else
          pK = Prefix & LCase(SplitWord(M, 1, "="))
          pV = ConvertProperty(SplitWord(M, 2, "=", , True))
On Error Resume Next
          Props.Add pV, pK
On Error GoTo 0
        End If
      Loop While True
      K = K + J - 1
      R = R & sSpace(I * SpIndent) & StartControl(L, Props, LMatch(M, "End"), CodeSection, Tag) & vbCrLf
      I = I + 1
      Stck(I) = Tag
    ElseIf L = "End" Then
      Set Props = Nothing
      Tag = Stck(I)
      I = I - 1
      If Tag <> "" Then
        R = R & sSpace(I * SpIndent) & EndControl(Tag) & vbCrLf
      End If
    End If
NextLine:
  Next
  ConvertFormUi = R
End Function

Private Function ConvertProperty(ByVal S As String) As String
  S = deQuote(S)
  S = DeComment(S)
  ConvertProperty = S
End Function

Private Function StartControl(ByVal L As String, ByVal Props As Collection, ByVal DoEmpty As Boolean, ByVal Code As String, ByRef TagType As String) As String
  Dim cType As String, cName As String, cIndex As String
  Dim tType As String, tCont As Boolean, tDef As String, Features As String
  Dim S As String, N As String, M As String
  Dim V As String
  N = vbCrLf
  TagType = ""
  
  cType = SplitWord(L, 2)
  cName = SplitWord(L, 3)
  cIndex = cValP(Props, "Index")
  If cIndex <> "" Then cName = cName & "_" & cIndex
  
  ControlData cType, tType, tCont, tDef, Features
  
  S = ""
On Error Resume Next
  If tType = "Line" Or tType = "Shape" Or tType = "Timer" Then
    Exit Function
  ElseIf tType = "Window" Then
    S = S & M & "<Window x:Class=""" & AssemblyName & ".Forms." & cName & """"
    S = S & N & "    xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation"""
    S = S & N & "    xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"""
    S = S & N & "    xmlns:d=""http://schemas.microsoft.com/expression/blend/2008"""
    S = S & N & "    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"""
    S = S & N & "    xmlns:local=""clr-namespace:" & AssemblyName & ".Forms"""
    S = S & N & "    mc:Ignorable=""d"""
    S = S & N & "    Title=" & Quote(cValP(Props, "caption"))
    S = S & M & "    Height=" & Quote(Px(cValP(Props, "clientheight", 0) + 435))
    S = S & M & "    Width=" & Quote(Px(cValP(Props, "clientwidth", 0) + 435))
    S = S & M & ">"
    S = S & N & " <Grid"
  ElseIf tType = "GroupBox" Then
    S = S & "<" & tType
    S = S & " x:Name=""" & cName & """"
    
    S = S & " Margin=" & Quote(Px(cValP(Props, "left")) & "," & Px(cValP(Props, "top")) & ",0,0")
    S = S & " Width=" & Quote(Px(cValP(Props, "width")))
    S = S & " Height=" & Quote(Px(cValP(Props, "height")))
    S = S & " VerticalAlignment=""Top"""
    S = S & " HorizontalAlignment=""Left"""
    S = S & " FontFamily=" & Quote(cValP(Props, "font.name", "Calibri"))
    S = S & " FontSize=" & Quote(cValP(Props, "font.size", 10))
      
    S = S & " Header=""" & cValP(Props, "caption") & """"
    S = S & "> <Grid Margin=""0,-15,0,0"""
  ElseIf tType = "Canvas" Then
    S = S & "<" & tType
    S = S & " x:Name=""" & cName & """"
    
    S = S & " Margin=" & Quote(Px(cValP(Props, "left")) & "," & Px(cValP(Props, "top")) & ",0,0")
    S = S & " Width=" & Quote(Px(cValP(Props, "width")))
    S = S & " Height=" & Quote(Px(cValP(Props, "height")))
  ElseIf tType = "Image" Then
    S = S & "<" & tType
    
    S = S & " x:Name=""" & cName & """"
    S = S & " Margin=" & Quote(Px(cValP(Props, "left")) & "," & Px(cValP(Props, "top")) & ",0,0")
    S = S & " Width=" & Quote(Px(cValP(Props, "width")))
    S = S & " Height=" & Quote(Px(cValP(Props, "height")))
    S = S & " VerticalAlignment=" & Quote("Top")
    S = S & " HorizontalAlignment=" & Quote("Left")
  Else
    S = ""
    S = S & "<" & tType
    S = S & " x:Name=""" & cName & """"
    S = S & " Margin=" & Quote(Px(cValP(Props, "left")) & "," & Px(cValP(Props, "top")) & ",0,0")
    S = S & " Padding=" & Quote("2,2,2,2")
    S = S & " Width=" & Quote(Px(cValP(Props, "width")))
    S = S & " Height=" & Quote(Px(cValP(Props, "height")))
    S = S & " VerticalAlignment=" & Quote("Top")
    S = S & " HorizontalAlignment=" & Quote("Left")
        
  End If

  If IsInStr(Features, "Font") Then
    S = S & " FontFamily=" & Quote(cValP(Props, "font.name", "Calibri"))
    S = S & " FontSize=" & Quote(cValP(Props, "font.size", 10))
    If Val(cValP(Props, "font.weight", "400")) > 400 Then S = S & " FontWeight=" & Quote("Bold")
      
  End If
  
  If IsInStr(Features, "Content") Then
    S = S & " Content=" & QuoteXML(cValP(Props, "caption") & cValP(Props, "text"))
  End If
  
  If IsInStr(Features, "Header") Then
    S = S & " Content=" & QuoteXML(cValP(Props, "caption") & cValP(Props, "text"))
  End If

  V = cValP(Props, "caption") & cValP(Props, "text")
  If IsInStr(Features, "Text") And V <> "" Then
    S = S & " Text=" & QuoteXML(V)
  End If
  
  V = cValP(Props, "ToolTipText")
  If IsInStr(Features, "ToolTip") And V <> "" Then
    S = S & " ToolTip=" & Quote(V)
  End If
  
  S = S & CheckControlEvents(tType, cName, Code)

  If DoEmpty Then
    S = S & " />"
    TagType = ""
  Else
    S = S & ">"
    TagType = tType
  End If
  StartControl = S
End Function

Public Function CheckControlEvents(ByVal ControlType As String, ByVal ControlName As String, Optional ByVal CodeSection As String) As String
  Dim Res As String
  Dim HasClick As Boolean, HasFocus As Boolean, HasChange As Boolean
  HasClick = True
  HasFocus = Not IsInStr("GroupBox", ControlType)
  HasChange = IsInStr("TextBox,ListBox", ControlType)
  
  Res = ""
  Res = Res & CheckEvent("MouseMove", ControlName, ControlType, CodeSection)
  If HasFocus Then
    Res = Res & CheckEvent("GotFocus", ControlName, ControlType, CodeSection)
    Res = Res & CheckEvent("LostFocus", ControlName, ControlType, CodeSection)
    Res = Res & CheckEvent("KeyDown", ControlName, ControlType, CodeSection)
    Res = Res & CheckEvent("KeyUp", ControlName, ControlType, CodeSection)
  End If
  If HasClick Then
    Res = Res & CheckEvent("Click", ControlName, ControlType, CodeSection)
    Res = Res & CheckEvent("DblClick", ControlName, ControlType, CodeSection)
  End If
  If HasChange Then
    Res = Res & CheckEvent("Change", ControlName, ControlType, CodeSection)
  End If

  CheckControlEvents = Res
End Function

Public Function CheckEvent(ByVal EventName As String, ByVal ControlName As String, ByVal ControlType As String, Optional ByVal CodeSection As String) As String
  Dim Search As String, Target As String, N As String
  Dim L As Long, V As String
  N = ControlName & "_" & EventName
  Search = " " & N & "("
  Target = EventName
  Select Case EventName
    Case "DblClick": Target = "MouseDoubleClick"
    Case "Change":
      If ControlType = "TextBox" Then Target = "TextChanged"
  End Select
  L = InStr(1, CodeSection, Search, vbTextCompare)
  If L > 0 Then
    V = Mid(CodeSection, L + 1, Len(N))   ' Get exact capitalization from source....
    CheckEvent = " " & Target & "=""" & V & """"
  Else
    CheckEvent = ""
  End If
End Function

Public Function EndControl(ByVal tType As String) As String
  Select Case tType
    Case "Line", "Shape", "Timer":
                          EndControl = ""
    Case "Window":        EndControl = " </Grid>" & vbCrLf & "</Window>"
    Case "GroupBox":      EndControl = "</Grid> </GroupBox>"
    Case Else:            EndControl = "</" & tType & ">"
  End Select
End Function


Public Function IsEvent(ByVal Str As String) As Boolean
  IsEvent = EventStub(Str) <> ""
End Function

Public Function EventStub(ByVal FName As String) As String
  Dim S As String, V As String, K As String
  S = "private void " & FName & "(object sender, RoutedEventArgs e) { "
  K = SplitWord(FName, 2, "_")
  Select Case K
    Case "Click", "DblClick", "Change", "Load", "GotFocus", "LostFocus"
      V = FName & "();"
    Case "QueryUnload"
      V = " long doCancel; long UnloadMode; " & FName & "(ref doCancel, ref UnloadMode);"
    Case "Validate", "Unload"
      V = "long doCancel; " & FName & "(ref doCancel);"
    Case "KeyDown", "KeyUp", "KeyPress"
    Case "MouseMove", "MouseDown", "MouseUp"
  End Select
  S = IIf(V = "", "", S & V & " }" & vbCrLf)
  
  EventStub = S
End Function
