Attribute VB_Name = "modSupportFiles"
Option Explicit

Public Function CreateProjectSupportFiles() As Boolean
  Dim S As String, F As String
  S = ApplicationXAML()
  F = "application.xaml"
  WriteOut F, S, ""
  
  S = VBExtensionClass()
  F = "VBExtension.cs"
  WriteOut F, S, ""
End Function

Public Function ApplicationXAML() As String
  Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "<Application x:Class=""Application"" "
  R = R & N & "xmlns = ""http://schemas.microsoft.com/winfx/2006/xaml/presentation"" "
  R = R & N & "xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"" "
  R = R & N & "xmlns:local=""clr-namespace:WpfApp1"" "
  R = R & N & "StartupUri=""MainWindow.xaml""> "
  R = R & N & "  <Application.Resources>"
  R = R & N & "  </Application.Resources>"
  R = R & N & "</Application>"

  ApplicationXAML = R
End Function



Public Function CreateProjectFile(ByVal vbpFile As String)
  Dim S As String, M As String, N As String
  Dim L
  S = ""
  M = ""
  N = vbCrLf
  

  S = S & M & "<?xml version=""1.0"" encoding=""utf-8""?>"
  S = S & N & "<Project ToolsVersion=""15.0"" xmlns=""http://schemas.microsoft.com/developer/msbuild/2003"">"
  S = S & N & "  <Import Project=""$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"" Condition=""Exists('$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props')"" />"
  S = S & N & "  <PropertyGroup>"
  S = S & N & "    <Configuration Condition="" '$(Configuration)' == '' "">Debug</Configuration>"
  S = S & N & "    <Platform Condition="" '$(Platform)' == '' "">AnyCPU</Platform>"
  S = S & N & "    <ProjectGuid>{92F75129-0EC1-47BA-85A7-E47F9EB140FD}</ProjectGuid>"
  S = S & N & "    <OutputType>WinExe</OutputType>"
  S = S & N & "    <RootNamespace>WpfApp2</RootNamespace>"
  S = S & N & "    <AssemblyName>WpfApp2</AssemblyName>"
  S = S & N & "    <TargetFrameworkVersion>v4.6.1</TargetFrameworkVersion>"
  S = S & N & "    <FileAlignment>512</FileAlignment>"
  S = S & N & "    <ProjectTypeGuids>{60dc8134-eba5-43b8-bcc9-bb4bc16c2548};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}</ProjectTypeGuids>"
  S = S & N & "    <WarningLevel>4</WarningLevel>"
  S = S & N & "    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>"
  S = S & N & "    <Deterministic>true</Deterministic>"
  S = S & N & "  </PropertyGroup>"
  S = S & N & "  <PropertyGroup Condition="" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' "">"
  S = S & N & "    <PlatformTarget>AnyCPU</PlatformTarget>"
  S = S & N & "    <DebugSymbols>true</DebugSymbols>"
  S = S & N & "    <DebugType>full</DebugType>"
  S = S & N & "    <Optimize>false</Optimize>"
  S = S & N & "    <OutputPath>bin\Debug\</OutputPath>"
  S = S & N & "    <DefineConstants>DEBUG;TRACE</DefineConstants>"
  S = S & N & "    <ErrorReport>prompt</ErrorReport>"
  S = S & N & "    <WarningLevel>4</WarningLevel>"
  S = S & N & "  </PropertyGroup>"
  S = S & N & "  <PropertyGroup Condition="" '$(Configuration)|$(Platform)' == 'Release|AnyCPU' "">"
  S = S & N & "    <PlatformTarget>AnyCPU</PlatformTarget>"
  S = S & N & "    <DebugType>pdbonly</DebugType>"
  S = S & N & "    <Optimize>true</Optimize>"
  S = S & N & "    <OutputPath>bin\Release\</OutputPath>"
  S = S & N & "    <DefineConstants>TRACE</DefineConstants>"
  S = S & N & "    <ErrorReport>prompt</ErrorReport>"
  S = S & N & "    <WarningLevel>4</WarningLevel>"
  S = S & N & "  </PropertyGroup>"
  S = S & N & "  <ItemGroup>"
  S = S & N & "    <Reference Include=""Microsoft.VisualBasic"" />"
  S = S & N & "    <Reference Include=""Microsoft.VisualBasic.Compatibility"" />"
  S = S & N & "    <Reference Include=""Microsoft.VisualBasic.Compatibility.Data"" />"
  S = S & N & "    <Reference Include=""System"" />"
  S = S & N & "    <Reference Include=""System.Data"" />"
  S = S & N & "    <Reference Include=""System.Xml"" />"
  S = S & N & "    <Reference Include=""Microsoft.CSharp"" />"
  S = S & N & "    <Reference Include=""System.Core"" />"
  S = S & N & "    <Reference Include=""System.Xml.Linq"" />"
  S = S & N & "    <Reference Include=""System.Data.DataSetExtensions"" />"
  S = S & N & "    <Reference Include=""System.Net.Http"" />"
  S = S & N & "    <Reference Include=""System.Xaml"">"
  S = S & N & "      <RequiredTargetFramework>4.0</RequiredTargetFramework>"
  S = S & N & "    </Reference>"
  S = S & N & "    <Reference Include=""WindowsBase"" />"
  S = S & N & "    <Reference Include=""PresentationCore"" />"
  S = S & N & "    <Reference Include=""PresentationFramework"" />"
  S = S & N & "  </ItemGroup>"
  S = S & N & "  <ItemGroup>"
  S = S & N & "    <ApplicationDefinition Include=""App.xaml"">"
  S = S & N & "      <Generator>MSBuild:Compile</Generator>"
  S = S & N & "      <SubType>Designer</SubType>"
  S = S & N & "    </ApplicationDefinition>"
  S = S & N & "    <Compile Include=""App.xaml.cs"">"
  S = S & N & "      <DependentUpon>App.xaml</DependentUpon>"
  S = S & N & "      <SubType>Code</SubType>"
  S = S & N & "    </Compile>"
  
  For Each L In Split(VBPForms(vbpFile), vbCrLf)
  S = S & N & "    <Page Include=""" & OutputSubFolder(L) & ChgExt(L, ".xaml") & """>"
  S = S & N & "      <SubType>Designer</SubType>"
  S = S & N & "      <Generator>MSBuild:Compile</Generator>"
  S = S & N & "    </Page>"
  S = S & N & "    <Compile Include=""" & OutputSubFolder(L) & ChgExt(L, ".xaml.cs") & """>"
  S = S & N & "      <DependentUpon>" & ChgExt(L, ".xaml") & "</DependentUpon>"
  S = S & N & "      <SubType>Code</SubType>"
  S = S & N & "    </Compile>"
  Next

  
  S = S & N & "    <Compile Include=""VBExtension.cs"" />"
  For Each L In Split(VBPClasses(vbpFile) & vbCrLf & VBPModules(vbpFile), vbCrLf)
  S = S & N & "    <Compile Include=""" & OutputSubFolder(L) & ChgExt(L, ".cs") & """ />"
  Next
  
  S = S & N & "  </ItemGroup>"
  S = S & N & "  <ItemGroup>"
  S = S & N & "    <Compile Include=""Properties\AssemblyInfo.cs"">"
  S = S & N & "      <SubType>Code</SubType>"
  S = S & N & "    </Compile>"
  S = S & N & "    <Compile Include=""Properties\Resources.Designer.cs"">"
  S = S & N & "      <AutoGen>True</AutoGen>"
  S = S & N & "      <DesignTime>True</DesignTime>"
  S = S & N & "      <DependentUpon>Resources.resx</DependentUpon>"
  S = S & N & "    </Compile>"
  S = S & N & "    <Compile Include=""Properties\Settings.Designer.cs"">"
  S = S & N & "      <AutoGen>True</AutoGen>"
  S = S & N & "      <DependentUpon>Settings.settings</DependentUpon>"
  S = S & N & "      <DesignTimeSharedInput>True</DesignTimeSharedInput>"
  S = S & N & "    </Compile>"
  S = S & N & "    <EmbeddedResource Include=""Properties\Resources.resx"">"
  S = S & N & "      <Generator>ResXFileCodeGenerator</Generator>"
  S = S & N & "      <LastGenOutput>Resources.Designer.cs</LastGenOutput>"
  S = S & N & "    </EmbeddedResource>"
  S = S & N & "    <None Include=""Properties\Settings.settings"">"
  S = S & N & "      <Generator>SettingsSingleFileGenerator</Generator>"
  S = S & N & "      <LastGenOutput>Settings.Designer.cs</LastGenOutput>"
  S = S & N & "    </None>"
  S = S & N & "  </ItemGroup>"
  S = S & N & "  <ItemGroup>"
  S = S & N & "    <None Include=""App.config"" />"
  S = S & N & "  </ItemGroup>"
  S = S & N & "  <Import Project=""$(MSBuildToolsPath)\Microsoft.CSharp.targets"" />"
  S = S & N & "</Project>"
  
  CreateProjectFile = S
  
  WriteOut ChgExt(FileName(vbpFile), ".csproj"), S
End Function

Public Function VBExtensionClass() As String
  Dim S As String, M As String, N As String
  Dim L
  S = ""
  M = ""
  N = vbCrLf
  
  S = S & M & ""
  S = S & M & UsingEverything
  S = S & N
  S = S & M & "public static class VBExtension {"
  S = S & N & "  public static object IIf(bool A, object B, object C) { return !!A ? B : C; }"
  S = S & N & "  public static bool IIf(bool A, bool B, bool C) { return !!A ? B : C; }"
  S = S & N & "  public static string IIf(bool A, string B, string C) { return !!A ? B : C; }"
  S = S & N & "  public static double IIf(bool A, double B, double C) { return !!A ? B : C; }"
  S = S & N & "  public static decimal IIf(bool A, decimal B, decimal C) { return !!A ? B : C; }"
  S = S & N & "  public static long IIf(bool A, long B, long C) { return !!A ? B : C; }"
  S = S & N
  S = S & N & "  public static bool IsMissing(object A) { return false; }"
  S = S & N & "  public static bool IsNull(object A) { return A == null; }"
  S = S & N
  S = S & N & "  public static System.DateTime NullDate() { try { return System.DateTime.Parse(""1/1/2001""); } catch { return Today; } }"
  S = S & N & "  public static bool IsDate(string D) { try { System.DateTime.Parse(D); } catch { return false; } return true; }"
  S = S & N
  S = S & N & "  public static System.DateTime CDate(object A) { return IsDate(A.ToString()) ? NullDate() : System.DateTime.Parse(A.ToString()); }"
  S = S & N & "  public static string CStr(object A) { return A.ToString(); }"
  S = S & N & "  public static bool CBool(object A) { return !!A; }"
  S = S & N
  S = S & N & "  public static System.DateTime DateValue(object A) { return CDate(A); }"
  S = S & M
  S = S & N & "}"
  
  
  VBExtensionClass = S
End Function
