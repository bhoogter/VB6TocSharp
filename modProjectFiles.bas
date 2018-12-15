Attribute VB_Name = "modProjectFiles"
Option Explicit

Public Function VBPModules(ByVal ProjectFile As String) As String
  Dim S As String, L
  Dim T As String
  Const C As String = "Module="
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
Public Function VBPForms(ByVal ProjectFile As String) As String
  Const WithExt As Boolean = True
  Dim S As String, L
  Dim T As String
  Const C As String = "Form="
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


Public Function VBPClasses(ByVal ProjectFile As String) As String
  Dim S As String, L
  Dim T As String
  Const C As String = "Class="
  S = ReadEntireFile(ProjectFile)
  For Each L In Split(S, vbCrLf)
    If Left(L, Len(C)) = C Then
      T = Mid(L, Len(C) + 1)
      If IsInStr(T, ";") Then T = SplitWord(T, 2, ";")
      VBPClasses = VBPClasses & IIf(VBPClasses = "", "", vbCrLf) & T
    End If
NextItem:
  Next
End Function

Public Function VBPUserControls(ByVal ProjectFile As String) As String
  Dim S As String, L
  Dim T As String
  Const C As String = "UserControl="
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


