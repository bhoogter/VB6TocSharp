Attribute VB_Name = "modProjectSpecific"
Option Explicit

Public Function ProjectSpecificPostCodeLineConvert(ByVal Str As String) As String
  Dim S As String
  S = Str
  
'  If IsInStr(S, "!C == null") Then Stop
  
  ' Some patterns we dont use or didn't catch in lint...
  If IsInStr(S, "DisposeDA") Then S = Replace(S, "DisposeDA", "// DisposeDA")
  If IsInStr(S, "MousePointer = vbNormal") Then S = Replace(S, "MousePointer = vbNormal", "MousePointer = vbDefault")

  ' We use decimal, not double
  If IsInStr(S, "Val(") Then S = Replace(S, "Val( ", "ValD(")
  
  ' Bad pattern combination
  If RegExTest(S, "\(!" & patToken & " == null\)") Then
    S = Replace(S, "!", "", 1)
    S = Replace(S, "==", "!=", 1)
  End If
  
  ' False ref entries...
  If IsInStr(S, "IsIn(") Then S = Replace(S, "ref ", "")
  If IsInStr(S, "POMode(") Then S = Replace(S, "ref ", "")
  If IsInStr(S, "OrderMode(") Then S = Replace(S, "ref ", "")
  If IsInStr(S, "InvenMode(") Then S = Replace(S, "ref ", "")
  
  ' Common Mistake Functions...
  If IsInStr(S, "StoreSettings.") Then S = Replace(S, "StoreSettings.", "StoreSettings().")
  
  ' etc
  If IsInStr(S, ".hwnd") Then S = Replace(S, ".hwnd", ".hWnd()")
  
  ProjectSpecificPostCodeLineConvert = S
End Function
