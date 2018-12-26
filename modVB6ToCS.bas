Attribute VB_Name = "modVB6ToCS"
Option Explicit

Public Function ConvertDefaultDefault(ByVal dType As String) As String
  Select Case dType
    Case "Long", "Double", "Currency", "Byte":
                      ConvertDefaultDefault = 0
    Case "Date":      ConvertDefaultDefault = """1/1/2001"""
    Case "String":    ConvertDefaultDefault = """"""
    Case "Boolean":   ConvertDefaultDefault = "false"
    Case Else:        ConvertDefaultDefault = "null"
  End Select
End Function

Public Function ConvertDataType(ByVal S As String) As String
  Select Case S
    Case "String":                ConvertDataType = "string"
    Case "Long":                  ConvertDataType = "long"
    Case "Double":                ConvertDataType = "double"
    Case "Variant":               ConvertDataType = "object"
    Case "Byte":                  ConvertDataType = "byte"
    Case "Boolean":               ConvertDataType = "bool"
    Case "Currency":              ConvertDataType = "decimal"
    Case "RecordSet", "ADODB.Recordset":
                                  ConvertDataType = "Recordset"
    Case "Connection", "ADODB.Connection":
                                  ConvertDataType = "Connection"
    Case "ADODB.Error":           ConvertDataType = "ADODB.Error"
    Case "ADODB.EventStatusEnum": ConvertDataType = "ADODB.EventStatusEnum"
    Case "Date":                  ConvertDataType = "System.DateTime"
    Case Else
      If IsInStr(VBPClasses(ClassNames:=True), S) Then
        ConvertDataType = S
      Else
        ConvertDataType = DefaultDataType
        Debug.Print "Unknown Data Type: " & S
      End If
  End Select
End Function

Public Function ControlData(ByVal cType As String, ByRef Name As String, ByRef Cont As Boolean, ByRef Def As String, ByRef Features As String)
  Cont = False
  Def = "Caption"
  Select Case cType
    Case "VB.Form":                   Name = "Window": Cont = True
    Case "VB.MDIForm":                Name = "Window": Cont = True
      Cont = True
    
    Case "VB.PictureBox":             Name = "Image": Cont = True: Def = "Picture": Features = "Tooltiptext"
    Case "VB.Label":                  Name = "Label": Features = "": Features = "Font,Content,Tooltiptext"
    Case "VB.TextBox":                Name = "TextBox": Def = "Text": Features = "Font,Text,Tooltiptext"
    Case "VB.Frame":                  Name = "GroupBox": Features = "Tooltiptext"
    Case "VB.CommandButton":          Name = "Button": Features = "Font,Content,Tooltiptext"
    Case "VB.CheckBox":               Name = "CheckBox": Features = "Font,Content,Tooltiptext"
    Case "VB.OptionButton":           Name = "RadioButton": Features = "Font,Content,Tooltiptext"
    Case "VB.ComboBox":               Name = "ComboBox": Def = "Text": Features = "Font,Text,Tooltiptext"
    Case "VB.ListBox":                Name = "ListBox": Def = "Text": Features = "Font,Tooltiptext"
    Case "VB.HScrollBar":             Name = "ScrollBar": Def = "Value": Features = ""
    Case "VB.VScrollBar":             Name = "ScrollBar": Def = "Value": Features = ""
    Case "VB.Timer":                  Name = "Timer": Def = "Enabled": Features = ""
    Case "VB.DriveListBox":           Name = "DriveListBox": Def = "Path": Features = ""
    Case "VB.DirListBox":             Name = "DirListBox": Def = "Path": Features = ""
    Case "VB.FileListBox":            Name = "FileListBox": Def = "Path": Features = ""
    Case "VB.Shape":                  Name = "Shape": Def = "Visible": Features = ""
    Case "VB.Line":                   Name = "Line": Def = "Visible": Features = ""
    Case "VB.Image":                  Name = "Image": Def = "Picture": Features = "Tooltiptext"
    Case "VB.Data":                   Name = "Data": Def = "DataSource": Features = ""
    Case "VB.OLE":                    Name = "OLE": Def = "OLE": Features = ""
    
    Case "VB.Menu":                   Name = "Menu"
    
    ' MS Windows Common Controls 6.0
    Case "MSComctlLib.TabStrip":
    Case "MSComctlLib.ToolBar":
    Case "MSComctlLib.StatusBar":       Name = "StatusBar": Def = "Text": Features = "Tooltiptext"
    Case "MSComctlLib.ProgressBar":     Name = "ProgressBar": Def = "Value": Features = "Tooltiptext"
    Case "MSComctlLib.TreeView":        Name = "TreeView": Features = "Tooltiptext"
    Case "MSComctlLib.ListView":        Name = "ListView": Features = "Tooltiptext"
    Case "MSComctlLib.ImageList":       Name = "ImageList": Features = "Tooltiptext"
    Case "MSComctlLib.Slider":          Name = "Slider"
    Case "MSComctlLib.ImageCombo":

    ' MS Windows Common Controls-2 6.0
'    Case "MSComCtl2.Animation":
    Case "MSComCtl2.UpDown":          Name = "Label"
    Case "MSComCtl2.DTPicker":        Name = "Label" '"DateTimePicker"
    Case "MSComCtl2.MonthView":       Name = "DatePicker"
    Case "MSComCtl2.FlatScrollBar":   Name = "ScrollBar"
    
    Case "MSComDlg.CommonDialog":     Name = "Label"
    Case "MSFlexGridLib.MSFlexGrid":  Name = "Grid"
    Case "MSDBGrid.DBGrid":           Name = "Grid"
    Case "TabDlg.SSTab":              Name = "TabControl"
    Case "RichTextLib.RichTextBox":   Name = "TextBlock"
    Case "InetCtlsObjects.Inet":      Name = "INet"
    Case "MSCommLib.MSComm":          Name = "MSComm"
    Case "MSWinsockLib.Winsock":      Name = "Winsock"
    
    Case "WinCDS.UGridIO":            Name = "UGridIO"
    Case "WinCDS.CandyButton":        Name = "Button"
    Case "WinCDS.ucPBar":             Name = "ProgressBar"
    Case "WinCDS.PrinterSelector":    Name = "Label"
    Case "WinCDS.RichTextBoxNew":     Name = "TextBlock"
    Case "WinCDS.MaskedPicture":      Name = "Image"
    
    Case "VJCZIPLib.VjcZip":          Name = "Label"
    Case "MSChart20Lib.MSChart":      Name = "Label"
    Case "MapPointCtl.MappointControl": Name = "Label"
    
    Case "LaVolpeAlphaImg.AlphaImgCtl": Name = "Image"
    Case "GIF89LibCtl.Gif89a":        Name = "Image"
    
    Case Else
      Debug.Print "Unknown Control Type: " & cType
      Name = "Label"
  End Select
End Function

 
Public Function ConvertVb6Specific(ByVal S As String, Optional ByRef Complete As Boolean)
  Dim W As String, R As String
  
  Complete = False
  W = SplitWord(Trim(S))
  R = SplitWord(Trim(S), 2, , , True)
  Select Case W
    Case "True": S = "true"
    Case "False": S = "false"
    Case "Nothing": S = "null"
    Case "vbTrue": S = "vbTriState.vbTrue"
    Case "vbFalse": S = "vbTriState.vbFalse"
    Case "vbUseDefault": S = "vbTriState.vbUseDefault"
    Case "Kill": S = "File.Delete(" & R & ")"
    Case "Open":    S = "VBOpenFile(" & Replace(SplitWord(R, 2, " As "), "#", "") & ", " & SplitWord(R, 1, " For ") & ")"
    Case "Print": S = "VBWriteFile(" & Replace(SplitWord(R, 1, ","), "#", "") & ", " & Replace(SplitWord(R, 2, ", ", , True), ";", ",") & ")"
    Case "Close": S = "VBCloseFile(" & Replace(R, "#", "") & ")"
    Case "New": S = "new " & R & "()": Complete = True
    Case "ReDim":
      Complete = True
      Dim RedimPres As Boolean, RedimVar As String, RedimTyp As String, RedimTmp As String, RedimMax As String, RedimIter As String
      If tLMatch(R, "Preserve ") Then
        R = Trim(tMid(R, 10))
        RedimPres = True
      End If
      
      RedimVar = RegExNMatch(R, patToken)
      RedimTyp = ConvertDataType(SubParam(RedimVar).asType)
      R = Trim(Replace(R, RedimVar, ""))
      If tLeft(R, 1) = "(" Then R = Mid(Trim(R), 2)
      RedimMax = Val(nextBy(R, ")"))
      RedimTmp = RedimVar & "_" & Random & "_tmp"
      RedimIter = "redim_iter_" & Random
      S = ""
      S = S & "List<" & RedimTyp & "> " & RedimTmp & " = new List<" & RedimTyp & ">();" & vbCrLf

      S = S & "for (int " & RedimIter & "=0;i<" & RedimMax & ";" & RedimIter & "++) {"
      If RedimPres Then
        S = S & RedimVar & ".Add(" & RedimIter & "<" & RedimVar & ".Count ? " & RedimVar & "(" & RedimIter & ") : " & ConvertDefaultDefault(SubParam(RedimVar).asType) & ");"
      Else
        S = S & RedimVar & ".Add(" & ConvertDefaultDefault(SubParam(RedimVar).asType) & ");"
      End If
      S = S & "}"
  End Select
  
  If IsInStr(S, ".Print ") Then
    If Right(S, 1) = ";" Then
      S = Replace(S, ".Print ", ".PrintNNL ")
      S = Left(S, Len(S) - 1)
    End If
    S = Replace(S, ";", ",")
  End If
  
  ConvertVb6Specific = S
End Function
