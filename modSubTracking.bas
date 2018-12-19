Attribute VB_Name = "modSubTracking"
Option Explicit

Private Type Variable
  Name As String
  asType As String
  asArray As String
  Param As Boolean
  Assigned As Boolean
  Used As Boolean
End Type

Private Vars() As Variable

Public Sub SubBegin()
  Dim nVars() As Variable
  Vars = nVars
End Sub

Private Function SubParamIndex(ByVal P As String) As Long
  On Error GoTo NoEntries
  For SubParamIndex = LBound(Vars) To UBound(Vars)
    If Vars(SubParamIndex).Name = P Then Exit Function
  Next
NoEntries:
  SubParamIndex = -1
End Function

Public Sub SubParamDecl(ByVal P As String, ByVal asType As String, ByVal asArray As String, ByVal isParam As Boolean)
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
  Dim K As Long
  K = SubParamIndex(P)
  If K >= 0 Then Vars(K).Assigned = True
End Sub

Public Sub SubParamUsed(ByVal P As String)
  Dim K As Long
  K = SubParamIndex(P)
  If K >= 0 Then Vars(K).Used = True
End Sub

Public Sub AddProperty(ByVal P As String, GSL As String, Body As String)

End Sub

Public Function ReadOutProperties() As String

End Function
