Attribute VB_Name = "modControlProperties"
Option Explicit

Public Function ConvertControlProperty(ByVal Src As String, ByVal vProp As String, ByVal cType As String) As String
'If IsInStr(vProp, "SetF") Then Stop
  ConvertControlProperty = vProp
  Select Case vProp
    Case "ListIndex": ConvertControlProperty = "SelectedIndex"
    Case "Visible": ConvertControlProperty = "Visibility"
    Case "Enabled": ConvertControlProperty = "IsEnabled"
    Case "TabStop": ConvertControlProperty = "IsTabStop"
    Case "SelStart": ConvertControlProperty = "SelectionStart"
    Case "SelLength": ConvertControlProperty = "SelectionLength"
    Case "Caption"
      If cType = "VB.Label" Then ConvertControlProperty = "Content"
    Case "Value"
      If cType = "VB.CheckBox" Then ConvertControlProperty = "IsChecked"
      If cType = "VB.RadioButton" Then ConvertControlProperty = "IsChecked"
      If cType = "MSComCtl2.DTPicker" Then ConvertControlProperty = "DisplayDate"
    Case ""
      If IsInStr(cType, "mage") Then Stop
      Select Case cType
        Case "VB.Caption":      ConvertControlProperty = "Content"
        Case "VB.TextBox":      ConvertControlProperty = "Text"
        Case "VB.ComboBox":     ConvertControlProperty = "Text"
        Case "VB.PictureBox":   ConvertControlProperty = "Source"
        Case "VB.ComboBox":     ConvertControlProperty = "Text"
        Case "VB.OptionButton": ConvertControlProperty = "IsChecked"
        Case "VB.CheckBox":     ConvertControlProperty = "IsChecked"
        Case "VB.Frame":        ConvertControlProperty = "Content"
        Case Else:              ConvertControlProperty = "DefaultProperty"
      End Select
  End Select
End Function
