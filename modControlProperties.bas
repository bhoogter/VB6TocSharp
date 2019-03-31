Attribute VB_Name = "modControlProperties"
Option Explicit

Public Function ConvertControlProperty(ByVal Src As String, ByVal vProp As String, ByVal cType As String) As String
  Select Case vProp
    Case "Visible": ConvertControlProperty = "Visibility"
    Case "Enabled": ConvertControlProperty = "IsEnabled"
    Case "Move": ConvertControlProperty = ""
    Case Else: ConvertControlProperty = vProp
  End Select
  Select Case cType
  End Select
End Function
