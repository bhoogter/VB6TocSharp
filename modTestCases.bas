Attribute VB_Name = "modTestCases"
Option Explicit

' This module exists solely to list test conversion caess to make sure the converter can convert itself containing them.
' There should be no active and/or used code in this module.
' These tests are not run, they are conversion tests.  They should be converted correctly when this project is converted.

Public Sub testCallModuleFunction()
' module name (w/, w/o)
' assign value (w/w/o)
' empty args parans (w/w/o)
  Dim S As String
  modGit.GitVersion
  GitVersion
  modGit.GitCmd "git --version"
  S = modGit.GitCmd("git --verison")
  GitCmd "git --version"
  S = GitCmd("git --verison")
  
  S = modGit.GitVersion()
  modGit.GitVersion
  S = GitVersion()
  GitVersion
End Sub

Public Sub testBooleans()
' not (w/w/o)
' if (w/w/o)
' fcall (w/w/o)
  Dim B As Boolean, NB As Boolean

  B = HasGit
  B = HasGit()
  B = modGit.HasGit
  B = modGit.HasGit()

  B = Not HasGit
  B = Not HasGit()
  B = Not modGit.HasGit
  B = Not modGit.HasGit()
  
  TestCallWithBooleanFunction HasGit
  TestCallWithBooleanFunction Not HasGit
  TestCallWithBooleanFunction modGit.HasGit
  TestCallWithBooleanFunction Not modGit.HasGit
  TestCallWithBooleanFunction HasGit()
  TestCallWithBooleanFunction Not HasGit()
  TestCallWithBooleanFunction modGit.HasGit()
  TestCallWithBooleanFunction Not modGit.HasGit()
  
  If HasGit Then Debug.Print ""
  If HasGit() Then Debug.Print ""
  If modGit.HasGit Then Debug.Print
  If modGit.HasGit() Then Debug.Print
  
  If Not HasGit Then Debug.Print ""
  If Not HasGit() Then Debug.Print ""
  If Not modGit.HasGit Then Debug.Print
  If Not modGit.HasGit() Then Debug.Print
End Sub

Public Function TestCallWithBooleanFunction(B As Boolean) As Boolean
  TestCallWithBooleanFunction = True
End Function

' Also have Property in a comment
Public Function testFunctionWithPropertyInName() As String()
  testFunctionWithPropertyInName = Array()
End Function


' This will only be readable if the file converts with correct braces.
Public Function TestFileFinishesWell() As Boolean
  TestFileFinishesWell = True
End Function
