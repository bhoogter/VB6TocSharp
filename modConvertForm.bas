Attribute VB_Name = "modConvertForm"
Option Explicit

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



Public Function ConvertFormUi(ByVal F As String) As String
  Dim Stck(0 To 100)
  Dim Sp, L, J As Long, K As Long, I As Long, Tag As String
  Dim M As String
  Dim R As String
  Dim Prefix As String
  Dim Props As Collection, pK As String, pV As String
  Sp = Split(F, vbCrLf)
  
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
            Prefix = Left(Prefix, InStrRev(Prefix, ".", -2))
          End If
        Else
          pK = Prefix & LCase(SplitWord(M, 1, "="))
          pV = ConvertProperty(SplitWord(M, 2, "=", , True))
          Props.Add pV, pK
        End If
      Loop While True
      K = K + J - 1
      R = R & sSpace(I * SpIndent) & StartControl(L, Props, LMatch(M, "End"), Tag) & vbCrLf
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

Private Function StartControl(ByVal L As String, ByVal Props As Collection, ByVal DoEmpty As Boolean, ByRef TagType As String) As String
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
    S = S & M & "<Window x:Class=""" & cName & """"
    S = S & N & "    xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation"""
    S = S & N & "    xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"""
    S = S & N & "    xmlns:d=""http://schemas.microsoft.com/expression/blend/2008"""
    S = S & N & "    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"""
    S = S & N & "    xmlns:local=""clr-namespace:WpfApp1"""
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
    S = S & " Content=" & Quote(cValP(Props, "caption") & cValP(Props, "text"))
  End If
  
  If IsInStr(Features, "Header") Then
    S = S & " Content=" & Quote(cValP(Props, "caption") & cValP(Props, "text"))
  End If

  V = cValP(Props, "caption") & cValP(Props, "text")
  If IsInStr(Features, "Text") And V <> "" Then
    S = S & " Text=" & Quote(V)
  End If
  
  V = cValP(Props, "ToolTipText")
  If IsInStr(Features, "ToolTip") And V <> "" Then
    S = S & " ToolTip=" & Quote(V)
  End If

  
  If DoEmpty Then
    S = S & " />"
    TagType = ""
  Else
    S = S & ">"
    TagType = tType
  End If
  StartControl = S
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

Public Function ControlData(ByVal cType As String, ByRef Name As String, ByRef Cont As Boolean, ByRef Def As String, ByRef Features As String)
  Cont = False
  Def = "Caption"
  Select Case cType
    Case "VB.Form"
      Name = "Window"
      Cont = True
    Case "VB.PictureBox":             Name = "Canvas": Cont = True: Def = "Picture": Features = "Tooltiptext"
    Case "VB.Label":                  Name = "Label": Features = "": Features = "Font,Content,Tooltiptext"
    Case "VB.TextBox":                Name = "TextBox": Def = "Text": Features = "Font,Text,Tooltiptext"
    Case "VB.Frame":                  Name = "GroupBox": Features = "Tooltiptext"
    Case "VB.CommandButton":          Name = "Button": Features = "Font,Content,Tooltiptext"
    Case "VB.CheckBox":               Name = "CheckBox": Features = "Font,Content,Tooltiptext"
    Case "VB.OptionButton":           Name = "RadioButton": Features = "Font,Content,Tooltiptext"
    Case "VB.ComboBox":               Name = "ComboBox": Def = "Text": Features = "Font,Text,Tooltiptext"
    Case "VB.ListBox":                Name = "ListBox": Def = "Text": Features = "Font,Tooltiptext"
    Case "VB.HScrollBar":             Name = "": Def = "Value": Features = ""
    Case "VB.VScrollBar":             Name = "": Def = "Value": Features = ""
    Case "VB.Timer":                  Name = "Timer": Def = "Enabled": Features = ""
    Case "VB.DriveListBox":           Name = "DriveListBox": Def = "Path": Features = ""
    Case "VB.DirListBox":             Name = "DirListBox": Def = "Path": Features = ""
    Case "VB.FileListBox":            Name = "FileListBox": Def = "Path": Features = ""
    Case "VB.Shape":                  Name = "Shape": Def = "Visible": Features = ""
    Case "VB.Line":                   Name = "Line": Def = "Visible": Features = ""
    Case "VB.Image":                  Name = "Canvas": Def = "Picture": Features = "Tooltiptext"
    Case "VB.Data":                   Name = "Data": Def = "DataSource": Features = ""
    Case "VB.OLE":                    Name = "OLE": Def = "OLE": Features = ""
    
    ' MS Windows Common Controls 6.0
    Case "MSComCtl2.TabStrip":
    Case "MSComCtl2.ToolBar":
    Case "MSComCtl2.StatusBar":       Name = "StatusBar": Def = "Text": Features = "Tooltiptext"
    Case "MSComCtl2.ProgressBar":     Name = "ProgressBar": Def = "Value": Features = "Tooltiptext"
    Case "MSComCtl2.TreeView":        Name = "TreeView": Features = "Tooltiptext"
    Case "MSComCtl2.ListView":        Name = "ListView": Features = "Tooltiptext"
    Case "MSComCtl2.ImageList":       Name = "ImageList": Features = "Tooltiptext"
    Case "MSComCtl2.Slider":
    Case "MSComCtl2.ImageCombo":

    ' MS Windows Common Controls-2 6.0
'    Case "MSComCtl2.Animation":
'    Case "MSComCtl2.UpDown":
    Case "MSComCtl2.DTPicker":        Name = "Label" '"DateTimePicker"
    Case "MSComCtl2.MonthView":       Name = "DatePicker"
    Case "MSComCtl2.FlatScrollBar":   Name = "ScrollBar"
    
    Case "MsFlexGridLib.MsFlexGrid":  Name = "Grid"
    Case "MSDBGrid.DBGrid":           Name = "Grid"
    Case "TabDlg.SSTab":              Name = "TabControl"
    Case "RichTextLib.RichTextBox":   Name = "TextBlock"
    Case "InetCtlsObjects.Inet":      Name = "INet"
    Case "MSCommLib.MSComm":          Name = "MSComm"
    Case "MSWinsockLib.Winsock":      Name = "Winsock"
    
    Case Else
      Debug.Print "Unknown Control Type: " & cType
      Name = "Label"
  End Select
End Function
