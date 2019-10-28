Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function INIWrite(ByVal sSection As String, ByVal sKeyName As String, ByVal sNewString As String, ByVal sINIFileName As String) As Boolean
On Error Resume Next
  WritePrivateProfileString sSection, sKeyName, sNewString, sINIFileName
  INIWrite = (Err.Number = 0)
End Function

Public Function INIRead(ByVal sSection As String, ByVal sKeyName As String, ByVal sINIFileName As String) As String
On Error Resume Next
  Dim sRet As String
  sRet = String(255, Chr(0))
  INIRead = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))
End Function

Public Function INISections(ByVal FileName As String) As String()
On Error Resume Next
  Dim strBuffer As String, intLen As Long

  Do While (intLen = Len(strBuffer) - 2) Or (intLen = 0)
    If strBuffer = vbNullString Then
      strBuffer = Space(256)
    Else
      strBuffer = String(Len(strBuffer) * 2, 0)
    End If
    
    intLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), FileName)
  Loop
  
  strBuffer = Left(strBuffer, intLen)
  INISections = Split(strBuffer, vbNullChar)
  ReDim Preserve INISections(UBound(INISections) - 1) As String
End Function

Public Function INISectionKeys(ByVal FileName As String, ByVal Section As String) As String()
On Error Resume Next
  Dim strBuffer As String, intLen As Long
  Dim I As Long, N As Long
  Dim RET() As String

  Do While (intLen = Len(strBuffer) - 2) Or (intLen = 0)
    If strBuffer = vbNullString Then
      strBuffer = Space(256)
    Else
      strBuffer = String(Len(strBuffer) * 2, 0)
    End If
    
    intLen = GetPrivateProfileSection(Section, strBuffer, Len(strBuffer), FileName)
    If intLen = 0 Then Exit Function
  Loop
  
  strBuffer = Left(strBuffer, intLen)
  RET = Split(strBuffer, vbNullChar)
  ReDim Preserve RET(UBound(RET) - 1) As String
  For I = LBound(RET) To UBound(RET)
    N = InStr(RET(I), "=")
    If N > 0 Then
      RET(I) = Left(RET(I), N - 1)
    Else
      Debug.Print "modINI.INISectionKeys - No '=' character found in line.  Section=" & Section & ", Line=" & RET(I) & ", file=" & FileName
    End If
  Next
  INISectionKeys = RET
End Function

Public Function ReadIniValue(ByVal INIPath As String, ByVal Key As String, ByVal Variable As String, Optional ByVal vDefault As String = "") As String
On Error Resume Next
  ReadIniValue = INIRead(Key, Variable, INIPath)
  If ReadIniValue = "" Then ReadIniValue = vDefault
End Function

Public Function WriteIniValue(ByVal INIPath As String, ByVal PutKey As String, ByVal PutVariable As String, ByVal PutValue As String, Optional ByVal DeleteOnEmpty As Boolean = False) As String
On Error Resume Next
  INIWrite PutKey, PutVariable, PutValue, INIPath
  WriteIniValue = INIRead(PutKey, PutVariable, INIPath)
End Function

