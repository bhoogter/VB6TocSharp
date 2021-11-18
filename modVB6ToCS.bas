Attribute VB_Name = "modVB6ToCS"
Option Explicit

Public Function ConvertDefaultDefault(ByVal DType As String) As String
  Select Case DType
    Case "Integer", "Long", "Double", "Currency", "Byte", "Single": ConvertDefaultDefault = 0
    Case "Date":      ConvertDefaultDefault = "DateTime.MinValue"
    Case "String":    ConvertDefaultDefault = """"""
    Case "Boolean":   ConvertDefaultDefault = "false"
    Case Else:        ConvertDefaultDefault = "null"
  End Select
End Function

Public Function ConvertDataType(ByVal S As String) As String
  Select Case S
    Case "Object", "Any", "Variant", "Variant()": ConvertDataType = DefaultDataType
    Case "Form", "Control":       ConvertDataType = "Window"
    Case "String":                ConvertDataType = "string"
    Case "String()":              ConvertDataType = "List<string>"
    Case "Long":                  ConvertDataType = "int"
    Case "Integer":               ConvertDataType = "int"
    Case "Double", "Single":      ConvertDataType = "decimal"
    Case "Variant":               ConvertDataType = "object"
    Case "Byte":                  ConvertDataType = "byte"
    Case "Boolean":               ConvertDataType = "bool"
    Case "Currency":              ConvertDataType = "decimal"
    Case "VbTriState":            ConvertDataType = "vbTriState"
    Case "Collection":            ConvertDataType = "Collection"
    Case "TSPNode":               ConvertDataType = "TSPNode"
    Case "TSPNetwork":            ConvertDataType = "TSPNetwork"
    Case "FindResults":           ConvertDataType = "FindResults"
    Case "Pushpin":               ConvertDataType = "Pushpin"
    Case "Map":                   ConvertDataType = "Map"
    Case "Node":                  ConvertDataType = "TreeViewItem"
    Case "Recordset", "ADODB.Recordset": ConvertDataType = "Recordset"
    Case "Connection", "ADODB.Connection": ConvertDataType = "Connection"
    Case "ADODB.Error":           ConvertDataType = "ADODB.Error"
    Case "ADODB.EventStatusEnum": ConvertDataType = "ADODB.EventStatusEnum"
    Case "SpeechLib.SpeechEngineConfidence", "SpeechLib.SpeechRecognitionType", "SpeechLib.ISpeechRecoResult", _
         "SpeechLib.SpeechInterference", "SpInprocRecognizer", "SpeechEngineConfidence", "ISpeechRecoGrammar", "SpSharedRecoContext"
      ConvertDataType = DefaultDataType
    
    Case "Date":                  ConvertDataType = "DateTime"
    Case "VbMsgBoxResult", "VbCompareMethod", "AlignConstants", _
         "stdole.IUnknown", "olelib.UUID", "olelib.STGMEDIUM", "olelib.FORMATETC", "olelib.BSCF", "olelib.IBinding", _
         "olelib.BINDINFO", "olelib.BINDF", "olelib.BINDSTATUS"
      ConvertDataType = S
    Case "XCTransaction2.XChargeTransaction", "PINPad"
      ConvertDataType = S
    Case "PictureBox", "Textbox", "Command", "ListBox", "ComboBox"
      ConvertDataType = S
    Case "MSCommLib.MSComm"
      ConvertDataType = S
            
    Case Else
      If IsInStr(VBPClasses(ClassNames:=True), S) Then
        ConvertDataType = S
      Else
        ConvertDataType = S
        Debug.Print "Unknown Data Type: " & S
      End If
  End Select
End Function

Public Sub ControlData(ByVal cType As String, ByRef Name As String, ByRef Cont As Boolean, ByRef Def As String, ByRef Features As String)
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
    Case "VB.DriveListBox":           Name = "usercontrols:DriveListBox": Def = "Path": Features = ""
    Case "VB.DirListBox":             Name = "usercontrols:DirListBox": Def = "Path": Features = ""
    Case "VB.FileListBox":            Name = "usercontrols:FileListBox": Def = "Path": Features = ""
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
    Case "MSComCtl2.UpDown":          Name = "usercontrols:UpDown"
    Case "MSComCtl2.DTPicker":        Name = "DatePicker"
    Case "MSComCtl2.MonthView":       Name = "DatePicker"
    Case "MSComCtl2.FlatScrollBar":   Name = "ScrollBar"
    
    Case "MSComDlg.CommonDialog":     Name = "Label"
    Case "MSFlexGridLib.MSFlexGrid":  Name = "usercontrols:FlexGrid"
    Case "MSDBGrid.DBGrid":           Name = "DataGrid"
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
End Sub

 
Public Function ConvertVb6Specific(ByVal S As String, Optional ByRef Complete As Boolean = False) As String
  Dim W As String, R As String
  
  Select Case Trim(S)
    Case "Array()":       S = "new List<dynamic>()"
    Case "App.Path":      S = "AppDomain.CurrentDomain.BaseDirectory"
  End Select
  
  Complete = False
  W = RegExNMatch(Trim(S), patToken)
  R = SplitWord(Trim(S), 2, , , True)
  Select Case W
    Case "True":          Complete = True: S = "true"
    Case "False":         Complete = True: S = "false"
    Case "Me":            Complete = True: S = "this"
    Case "Nothing":       Complete = True: S = "null"
    Case "vbTrue":        Complete = True: S = "vbTriState.vbTrue"
    Case "vbFalse":       Complete = True: S = "vbTriState.vbFalse"
    Case "vbUseDefault":  Complete = True: S = "vbTriState.vbUseDefault"
    Case "Date", "Today": Complete = True: S = "DateTime.Today;"
    Case "Now":           Complete = True: S = "DateTime.Now;"
    Case "Kill":          S = "File.Delete(" & R & ");"
    Case "FreeFile":      S = "FreeFile();"
    Case "Open":          S = "VBOpenFile(" & Replace(SplitWord(R, 2, " As "), "#", "") & ", " & SplitWord(R, 1, " For ") & ");"
    Case "Print":         S = "VBWriteFile(" & Replace(SplitWord(R, 1, ","), "#", "") & ", " & Replace(SplitWord(R, 2, ", ", , True), ";", ",") & ");"
    Case "Close":         S = "VBCloseFile(" & Replace(R, "#", "") & ");"
    Case "New":           Complete = True: S = "new " & R & "();"
    Case "vbAlignLeft":   S = "AlignConstants.vbAlignLeft"
    Case "vbAlignRight":  S = "AlignConstants.vbAlignRight"
    Case "vbAlignTop":    S = "AlignConstants.vbAlignTop"
    Case "vbAlignBottom": S = "AlignConstants.vbAlignBottom"
    Case "RaiseEvent":
      Complete = True
      W = RegExNMatch(R, patToken)
      R = Mid(R, Len(W) + 1)
      If R = "" Then R = "()"
      S = "event" & W & "?.Invoke" & R & ";"
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

Public Function ConvertVb6Syntax(ByVal S As String) As String
  Dim W As String, R As String
  W = RegExNMatch(Trim(S), patToken)
  R = SplitWord(Trim(S), 2, , , True)
  Select Case W
    Case "Open":          S = "VBOpenFile(" & Replace(SplitWord(R, 2, " As "), "#", "") & ", " & SplitWord(R, 1, " For ") & ")"
    Case "Print":         S = "VBWriteFile(" & Replace(SplitWord(R, 1, ","), "#", "") & ", " & Replace(SplitWord(R, 2, ", ", , True), ";", ",") & ")"
    Case "Input":         S = "VBReadFile(" & Replace(SplitWord(R, 1, ","), "#", "") & ", " & Replace(SplitWord(R, 2, ", ", , True), ";", ",") & ")"
    Case "Line":          S = "VBReadFileLine(" & Replace(SplitWord(R, 1, ","), "#", "") & ", " & Replace(SplitWord(R, 2, ", ", , True), ";", ",") & ")"
    Case "Close":         S = "VBCloseFile(" & Replace(R, "#", "") & ")"
    Case "New":           S = "new " & R & "()"
    Case "RaiseEvent":
      W = RegExNMatch(R, patToken)
      R = Mid(R, Len(W) + 1)
      If R = "" Then R = "()"
      S = "event" & W & "?.Invoke" & R
  End Select
  
  ConvertVb6Syntax = S
End Function
