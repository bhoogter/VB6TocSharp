Attribute VB_Name = "modConvert"
Option Explicit

Const WithMark As String = "_WithVar_"

Dim WithLevel As Long, MaxWithLevel As Long
Dim WithVars As String, WithTypes As String, WithAssign As String
Dim FormName As String

Dim CurrentModule As String

Dim CurrSub As String

Public Const CONVERTER_VERSION_1 As String = "v1"
Public Const CONVERTER_VERSION_2 As String = "v2"
Public Const CONVERTER_VERSION_DEFAULT As String = CONVERTER_VERSION_2

Public Function QuickConvertProject() As Boolean
  QuickConvertProject = ConvertProject(vbpFile, CONVERTER_VERSION_2)
End Function

Public Function QuickConvert() As Boolean
  QuickConvert = ConvertFile("modQuickConvert.bas", False, CONVERTER_VERSION_2)
End Function

Public Function ConvertProject(Optional ByVal vbpFile As String = "", Optional ByVal ConverterVersion As String = CONVERTER_VERSION_DEFAULT) As Boolean
  If vbpFile = "" Then vbpFile = modConfig.vbpFile
  Prg 0, 1, "Preparing..."
  ScanRefs
  CreateProjectFile vbpFile
  CreateProjectSupportFiles
  ConvertFileList FilePath(vbpFile), VBPModules(vbpFile) & vbCrLf & VBPClasses(vbpFile) & vbCrLf & VBPForms(vbpFile), vbCrLf, ConverterVersion
  ConvertProject = True
End Function

Public Function ConvertFileList(ByVal Path As String, ByVal List As String, Optional ByVal Sep As String = vbCrLf, Optional ByVal ConverterVersion As String = CONVERTER_VERSION_DEFAULT) As Boolean
  Dim L As Variant, V As Long, N As Long
  V = StrCnt(List, Sep) + 1
  Prg 0, V, N & "/" & V & "..."
  For Each L In Split(List, Sep)
    N = N + 1
    If L = "" Then GoTo NextItem
    
    If L = "modFunctionList.bas" Then GoTo NextItem
    
    ConvertFile Path & L, False, ConverterVersion
    
NextItem:
    Prg N, , N & "/" & V & ": " & L
    DoEvents
  Next
  Prg
End Function

Public Function ConvertFile(ByVal SomeFile As String, Optional ByVal UIOnly As Boolean = False, Optional ByVal ConverterVersion As String = CONVERTER_VERSION_DEFAULT) As Boolean
  If Not IsInStr(SomeFile, "\") Then SomeFile = vbpPath & SomeFile
  CurrentModule = ""
  Select Case LCase(FileExt(SomeFile))
    Case ".bas": ConvertFile = ConvertModule(SomeFile, ConverterVersion)
    Case ".cls": ConvertFile = ConvertClass(SomeFile, ConverterVersion)
    Case ".frm": FormName = FileBaseName(SomeFile): ConvertFile = ConvertForm(SomeFile, UIOnly, ConverterVersion)
'      Case ".ctl": ConvertModule  someFile
    Case Else: MsgBox "UNKNOWN VB TYPE: " & SomeFile: Exit Function
  End Select
  FormName = ""
  ConvertFile = True
End Function

Public Function ConvertForm(ByVal frmFile As String, Optional ByVal UIOnly As Boolean = False, Optional ByVal ConverterVersion As String = CONVERTER_VERSION_DEFAULT) As Boolean
  Dim S As String, J As Long, Preamble As String, Code As String, Globals As String, Functions As String
  Dim X As String, fName As String
  Dim F As String
  If Not FileExists(frmFile) Then
    MsgBox "File not found in ConvertForm: " & frmFile
    Exit Function
  End If
  
  S = ReadEntireFile(frmFile)
  fName = ModuleName(S)
  CurrentModule = fName
  F = fName & ".xaml.cs"
  If IsConverted(F, frmFile) Then Debug.Print "Form Already Converted: " & F: Exit Function
  
  J = CodeSectionLoc(S)
  Preamble = Left(S, J - 1)
  Code = Mid(S, J)
  
  X = ConvertFormUi(Preamble, Code)
  F = fName & ".xaml"
  WriteOut F, X, frmFile
  If UIOnly Then Exit Function
  
  
  Dim ConvertedCode As String
  If ConverterVersion = CONVERTER_VERSION_2 Then
    ConvertedCode = ""
    Dim ControlArrays As String, VV As Variant
    ControlArrays = Replace(Replace(Replace(modConvertForm.FormControlArrays, "][", ";"), "[", ""), "]", "")
    For Each VV In Split(ControlArrays, ";")
      Dim ControlArrayParts() As String
      ControlArrayParts = Split(VV, ",")
      ConvertedCode = ConvertedCode & "public List<" & ControlArrayParts(1) & "> " & ControlArrayParts(0) & " { get => VBExtension.controlArray<" & ControlArrayParts(1) & ">(this, """ & ControlArrayParts(0) & """); }" & vbCrLf2

'      ConvertedCode = ConvertedCode & "public ControlArrayList<" & ControlArrayParts(1) & "> " & ControlArrayParts(0) & "() => VBExtension.controlArray(this, """ & ControlArrayParts(1) & """).Cast<" & ControlArrayParts(1) & ">().ToList();" & vbCrLf
'      ConvertedCode = ConvertedCode & "public " & ControlArrayParts(1) & " " & ControlArrayParts(0) & "(int i) => " & ControlArrayParts(0) & "()[i];" & vbCrLf2
    Next
  
    ConvertedCode = ConvertedCode & QuickConvertFile(frmFile)
  Else
    J = CodeSectionGlobalEndLoc(Code)
    Globals = ConvertGlobals(Left(Code, J))
    InitLocalFuncs FormControls(fName, Preamble) & ScanRefsFileToString(frmFile)
    Functions = ConvertCodeSegment(Mid(Code, J))
    ConvertedCode = Globals & vbCrLf2 & Functions
  End If
  
  X = ""
  X = X & UsingEverything(fName) & vbCrLf
  X = X & vbCrLf
  X = X & "namespace " & AssemblyName & ".Forms" & vbCrLf
  X = X & "{" & vbCrLf
  X = X & "public partial class " & fName & " : Window {" & vbCrLf
  X = X & "  private static " & fName & " _instance;" & vbCrLf
  X = X & "  public static " & fName & " instance { set { _instance = null; } get { return _instance ?? (_instance = new " & fName & "()); }}"
  X = X & "  public static void Load() { if (_instance == null) { dynamic A = " + fName + ".instance; } }"
  X = X & "  public static void Unload() { if (_instance != null) instance.Close(); _instance = null; }"
  X = X & "  public " & fName & "() => InitializeComponent();" & vbCrLf
  X = X & vbCrLf
  X = X & vbCrLf
  X = X & ConvertedCode
  X = X & vbCrLf & "}"
  X = X & vbCrLf & "}"
  
  X = deWS(X)
  
  F = fName & ".xaml.cs"
  WriteOut F, X, frmFile
End Function

Public Function ConvertModule(ByVal basFile As String, Optional ByVal ConverterVersion As String = CONVERTER_VERSION_DEFAULT) As Boolean
  Dim S As String, J As Long, Code As String, Globals As String, Functions As String
  Dim F As String, X As String, fName As String
  If Not FileExists(basFile) Then
    MsgBox "File not found in ConvertModule: " & basFile
    Exit Function
  End If
  S = ReadEntireFile(basFile)
  fName = ModuleName(S)
  CurrentModule = fName
  F = fName & ".cs"
  If IsConverted(F, basFile) Then Debug.Print "Module Already Converted: " & F: Exit Function
  
  fName = ModuleName(S)
  Code = Mid(S, CodeSectionLoc(S))
  
  Dim UserCode As String
  If ConverterVersion = CONVERTER_VERSION_2 Then
    UserCode = QuickConvertFile(basFile)
  Else
    J = CodeSectionGlobalEndLoc(Code)
    Globals = ConvertGlobals(Left(Code, J - 1), True)
    Functions = ConvertCodeSegment(Mid(Code, J), True)
    
    UserCode = nlTrim(Globals & vbCrLf & vbCrLf & Functions)
    UserCode = deWS(UserCode)
  End If
  
  X = ""
  X = X & UsingEverything(fName) & vbCrLf
  X = X & vbCrLf
  X = X & "static class " & fName & " {" & vbCrLf
  X = X & UserCode
  X = X & vbCrLf & "}"
  
  WriteOut F, X, basFile
End Function

Public Function ConvertClass(ByVal clsFile As String, Optional ByVal ConverterVersion As String = CONVERTER_VERSION_DEFAULT) As Boolean
  Dim S As String, J As Long, Code As String, Globals As String, Functions As String
  Dim F As String, X As String, fName As String
  Dim cName As String
  If Not FileExists(clsFile) Then
    MsgBox "File not found in ConvertModule: " & clsFile
    Exit Function
  End If
  S = ReadEntireFile(clsFile)
  fName = ModuleName(S)
  CurrentModule = fName
  F = fName & ".cs"
  If IsConverted(F, clsFile) Then Debug.Print "Class Already Converted: " & F: Exit Function

  
  Dim UserCode As String
  If ConverterVersion = CONVERTER_VERSION_2 Then
    UserCode = QuickConvertFile(clsFile)
  Else
    Code = Mid(S, CodeSectionLoc(S))
    
    J = CodeSectionGlobalEndLoc(Code)
    Globals = ConvertGlobals(Left(Code, J - 1))
    Functions = ConvertCodeSegment(Mid(Code, J))
    
    UserCode = deWS(Globals & vbCrLf & vbCrLf & Functions)
  End If
  
  X = ""
  X = X & UsingEverything(fName) & vbCrLf
  X = X & vbCrLf
  X = X & "public class " & fName & " {" & vbCrLf
  X = X & UserCode
  X = X & vbCrLf & "}"
  
  F = fName & ".cs"
  WriteOut F, X, clsFile
End Function
