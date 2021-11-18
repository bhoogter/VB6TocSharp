Attribute VB_Name = "modProjectFiles"
Option Explicit

Public Function VBPCode(Optional ByVal ProjectFile As String = "") As String
  VBPCode = VBPModules & vbCrLf & VBPForms & vbCrLf & VBPClasses & vbCrLf & VBPUserControls
End Function

Public Function VBPModules(Optional ByVal ProjectFile As String = "") As String
  Dim S As String, L As Variant
  Dim T As String
  Const C As String = "Module="
  If ProjectFile = "" Then ProjectFile = vbpFile
  S = ReadEntireFile(ProjectFile)
  For Each L In Split(S, vbCrLf)
    If Left(L, Len(C)) = C Then
      T = Mid(L, Len(C) + 1)
      If IsInStr(T, ";") Then T = SplitWord(T, 2, ";")
'If IsInStr(LCase(T), "subclass") Then Stop
      If LCase(T) = "modlistsubclass.bas" Then GoTo NextItem
      VBPModules = VBPModules & IIf(VBPModules = "", "", vbCrLf) & T
    End If
NextItem:
  Next
End Function
Public Function VBPForms(Optional ByVal ProjectFile As String = "") As String
  Const WithExt As Boolean = True
  Dim S As String, L As Variant
  Dim T As String
  Const C As String = "Form="
  If ProjectFile = "" Then ProjectFile = vbpFile
  S = ReadEntireFile(ProjectFile)
  For Each L In Split(S, vbCrLf)
    If Left(L, Len(C)) = C Then
      T = Mid(L, Len(C) + 1)
      If IsInStr(T, ";") Then T = SplitWord(T, 1, ";")
      If Not WithExt And Right(T, 4) = ".frm" Then T = Left(T, Len(T) - 4)
      Select Case LCase(T)
        Case "faxtest":           T = "FaxPO"
        Case "frmpos":            T = "frmCashRegister"
        Case "frmposquantity":    T = "frmCashRegisterQuantity"
        Case "calendarinst":      T = "CalendarInstr"
        Case "frmedi":            T = "frmAshleyEDIItemAlign"
        Case "frmpracticefiles":  T = "PracticeFiles"
        Case "txttextselect":     T = "frmSelectText"
      End Select
      VBPForms = VBPForms & IIf(VBPForms = "", "", vbCrLf) & T
    End If
NextItem:
  Next
End Function

Public Function VBPClasses(Optional ByVal ProjectFile As String = "", Optional ByVal ClassNames As Boolean = False) As String
  Dim S As String, L As Variant
  Dim T As String
  Const C As String = "Class="
  If ProjectFile = "" Then ProjectFile = vbpFile
  S = ReadEntireFile(ProjectFile)
  For Each L In Split(S, vbCrLf)
    If Left(L, Len(C)) = C Then
      T = Mid(L, Len(C) + 1)
      If IsInStr(T, ";") Then T = SplitWord(T, 2, ";")
      VBPClasses = VBPClasses & IIf(VBPClasses = "", "", vbCrLf) & T
    End If
NextItem:
  Next
  If ClassNames Then VBPClasses = Replace(VBPClasses, ".cls", "")
End Function

Public Function VBPUserControls(Optional ByVal ProjectFile As String = "") As String
  Dim S As String, L As Variant
  Dim T As String
  Const C As String = "UserControl="
  If ProjectFile = "" Then ProjectFile = vbpFile
  S = ReadEntireFile(ProjectFile)
  For Each L In Split(S, vbCrLf)
    If Left(L, Len(C)) = C Then
      T = Mid(L, Len(C) + 1)
      If IsInStr(T, ";") Then T = SplitWord(T, 2, ";")
      VBPUserControls = VBPUserControls & IIf(VBPUserControls = "", "", vbCrLf) & T
    End If
NextItem:
  Next
End Function
