using VB6 = Microsoft.VisualBasic.Compatibility.VB6;
using System.Runtime.InteropServices;
using static VBExtension;
using static VBConstants;
using Microsoft.VisualBasic;
using System;
using System.Windows;
using System.Windows.Controls;
using static System.DateTime;
using static System.Math;
using static Microsoft.VisualBasic.Globals;
using static Microsoft.VisualBasic.Collection;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Conversion;
using static Microsoft.VisualBasic.DateAndTime;
using static Microsoft.VisualBasic.ErrObject;
using static Microsoft.VisualBasic.FileSystem;
using static Microsoft.VisualBasic.Financial;
using static Microsoft.VisualBasic.Information;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static Microsoft.VisualBasic.VBMath;
using System.Collections.Generic;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ColorConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.DrawStyleConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.FillStyleConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.GlobalModule;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.Printer;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterCollection;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterObjectConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ScaleModeConstants;
using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.SystemColorConstants;
using ADODB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using VB2CS.Forms;
using static modUtils;
using static modConvert;
using static modProjectFiles;
using static modTextFiles;
using static modRegEx;
using static frmTest;
using static modConvertForm;
using static modSubTracking;
using static modVB6ToCS;
using static modUsingEverything;
using static modSupportFiles;
using static modConfig;
using static modRefScan;
using static modConvertUtils;
using static modControlProperties;
using static modProjectSpecific;
using static modINI;
using static modLinter;
using static modGit;
using static modDirStack;
using static modShell;
using static VB2CS.Forms.frm;
using static VB2CS.Forms.frmConfig;


static class modSupportFiles {
// Option Explicit


public static bool CreateProjectSupportFiles() {
  bool CreateProjectSupportFiles = false;
  string S = "";
  string F = "";

  S = ApplicationXAML();
  F = "application.xaml";
  WriteOut(F, S, "");

  S = VBExtensionClass();
  F = "VBExtension.cs";
  WriteOut(F, S, "");

  S = VBAConstantsClass();
  F = "VBConstants.cs";
  WriteOut(F, S, "");
  return CreateProjectSupportFiles;
}

public static string ApplicationXAML() {
  string ApplicationXAML = "";
  string R = "";
  string M = "";
  string N = "";

  R = "";
  M = "";
  N = vbCrLf;

  R = R + M + "<Application x:Class=\"Application\" ";
  R = R + N + "xmlns = \"http://schemas.microsoft.com/winfx/2006/xaml/presentation\" ";
  R = R + N + "xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\" ";
  R = R + N + "xmlns:local=\"clr-namespace:WpfApp1\" ";
  R = R + N + "StartupUri=\"MainWindow.xaml\"> ";
  R = R + N + "  <Application.Resources>";
  R = R + N + "  </Application.Resources>";
  R = R + N + "</Application>";

  ApplicationXAML = R;
  return ApplicationXAML;
}

public static dynamic CreateProjectFile(string vbpFile) {
  dynamic CreateProjectFile = null;
  string S = "";
  string M = "";
  string N = "";

  dynamic L = null;

  S = "";
  M = "";
  N = vbCrLf;


  S = S + M + "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
  S = S + N + "<Project ToolsVersion=\"15.0\" xmlns=\"http://schemas.microsoft.com/developer/msbuild/2003\">";
  S = S + N + "  <Import Project=\"$(MSBuildExtensionsPath)\\$(MSBuildToolsVersion)\\Microsoft.Common.props\" Condition=\"Exists('$(MSBuildExtensionsPath)\\$(MSBuildToolsVersion)\\Microsoft.Common.props')\" />";
  S = S + N + "  <PropertyGroup>";
  S = S + N + "    <Configuration Condition=\" '$(Configuration)' == '' \">Debug</Configuration>";
  S = S + N + "    <Platform Condition=\" '$(Platform)' == '' \">AnyCPU</Platform>";
  S = S + N + "    <ProjectGuid>{92F75129-0EC1-47BA-85A7-E47F9EB140FD}</ProjectGuid>";
  S = S + N + "    <OutputType>WinExe</OutputType>";
  S = S + N + "    <RootNamespace>" + AssemblyName() + "</RootNamespace>";
  S = S + N + "    <AssemblyName>" + AssemblyName() + "</AssemblyName>";
  S = S + N + "    <TargetFrameworkVersion>v4.6.1</TargetFrameworkVersion>";
  S = S + N + "    <FileAlignment>512</FileAlignment>";
  S = S + N + "    <ProjectTypeGuids>{60dc8134-eba5-43b8-bcc9-bb4bc16c2548};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}</ProjectTypeGuids>";
  S = S + N + "    <WarningLevel>4</WarningLevel>";
  S = S + N + "    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>";
  S = S + N + "    <Deterministic>true</Deterministic>";
  S = S + N + "  </PropertyGroup>";
  S = S + N + "  <PropertyGroup Condition=\" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' \">";
  S = S + N + "    <PlatformTarget>AnyCPU</PlatformTarget>";
  S = S + N + "    <DebugSymbols>true</DebugSymbols>";
  S = S + N + "    <DebugType>full</DebugType>";
  S = S + N + "    <Optimize>false</Optimize>";
  S = S + N + "    <OutputPath>bin\\Debug\\</OutputPath>";
  S = S + N + "    <DefineConstants>DEBUG;TRACE</DefineConstants>";
  S = S + N + "    <ErrorReport>prompt</ErrorReport>";
  S = S + N + "    <WarningLevel>4</WarningLevel>";
  S = S + N + "  </PropertyGroup>";
  S = S + N + "  <PropertyGroup Condition=\" '$(Configuration)|$(Platform)' == 'Release|AnyCPU' \">";
  S = S + N + "    <PlatformTarget>AnyCPU</PlatformTarget>";
  S = S + N + "    <DebugType>pdbonly</DebugType>";
  S = S + N + "    <Optimize>true</Optimize>";
  S = S + N + "    <OutputPath>bin\\Release\\</OutputPath>";
  S = S + N + "    <DefineConstants>TRACE</DefineConstants>";
  S = S + N + "    <ErrorReport>prompt</ErrorReport>";
  S = S + N + "    <WarningLevel>4</WarningLevel>";
  S = S + N + "  </PropertyGroup>";
  S = S + N + "  <ItemGroup>";
  S = S + N + "    <Reference Include=\"Microsoft.VisualBasic\" />";
  S = S + N + "    <Reference Include=\"Microsoft.VisualBasic.Compatibility\" />";
  S = S + N + "    <Reference Include=\"Microsoft.VisualBasic.Compatibility.Data\" />";
  S = S + N + "    <Reference Include=\"Microsoft.VisualBasic.PowerPacks, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL\" />";
  S = S + N + "    <Reference Include=\"Microsoft.VisualBasic.PowerPacks.Vs, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL\" />";
  S = S + N + "    <Reference Include=\"System\" />";
  S = S + N + "    <Reference Include=\"System.Data\" />";
  S = S + N + "    <Reference Include=\"System.Xml\" />";
  S = S + N + "    <Reference Include=\"Microsoft.CSharp\" />";
  S = S + N + "    <Reference Include=\"System.Core\" />";
  S = S + N + "    <Reference Include=\"System.Xml.Linq\" />";
  S = S + N + "    <Reference Include=\"System.Data.DataSetExtensions\" />";
  S = S + N + "    <Reference Include=\"System.Net.Http\" />";
  S = S + N + "    <Reference Include=\"System.Xaml\">";
  S = S + N + "      <RequiredTargetFramework>4.0</RequiredTargetFramework>";
  S = S + N + "    </Reference>";
  S = S + N + "    <Reference Include=\"WindowsBase\" />";
  S = S + N + "    <Reference Include=\"PresentationCore\" />";
  S = S + N + "    <Reference Include=\"PresentationFramework\" />";
  S = S + N + "  </ItemGroup>";
  S = S + N + "  <ItemGroup>";
  S = S + N + "    <ApplicationDefinition Include=\"App.xaml\">";
  S = S + N + "      <Generator>MSBuild:Compile</Generator>";
  S = S + N + "      <SubType>Designer</SubType>";
  S = S + N + "    </ApplicationDefinition>";
  S = S + N + "    <Compile Include=\"App.xaml.cs\">";
  S = S + N + "      <DependentUpon>App.xaml</DependentUpon>";
  S = S + N + "      <SubType>Code</SubType>";
  S = S + N + "    </Compile>";

  foreach(var L in Split(VBPForms(vbpFile), vbCrLf)) {
    if (L == "") {
goto SkipForm;
    }
    S = S + N + "    <Page Include=\"" + OutputSubFolder(L) + ChgExt(L, ".xaml") + "\">";
    S = S + N + "      <SubType>Designer</SubType>";
    S = S + N + "      <Generator>MSBuild:Compile</Generator>";
    S = S + N + "    </Page>";
    S = S + N + "    <Compile Include=\"" + OutputSubFolder(L) + ChgExt(L, ".xaml.cs") + "\">";
    S = S + N + "      <DependentUpon>" + ChgExt(L, ".xaml") + "</DependentUpon>";
    S = S + N + "      <SubType>Code</SubType>";
    S = S + N + "    </Compile>";
SkipForm:
  }


  S = S + N + "    <Compile Include=\"VBExtension.cs\" />";
  S = S + N + "    <Compile Include=\"VBConstants.cs\" />";
  foreach(var L in Split(VBPClasses(vbpFile) & vbCrLf & VBPModules(vbpFile), vbCrLf)) {
    if (L == "") {
goto SkipClass;
    }
    S = S + N + "    <Compile Include=\"" + OutputSubFolder(L) + ChgExt(L, ".cs") + "\" />";
SkipClass:
  }

  S = S + N + "  </ItemGroup>";
  S = S + N + "  <ItemGroup>";
  S = S + N + "    <Compile Include=\"Properties\\AssemblyInfo.cs\">";
  S = S + N + "      <SubType>Code</SubType>";
  S = S + N + "    </Compile>";
  S = S + N + "    <Compile Include=\"Properties\\Resources.Designer.cs\">";
  S = S + N + "      <AutoGen>True</AutoGen>";
  S = S + N + "      <DesignTime>True</DesignTime>";
  S = S + N + "      <DependentUpon>Resources.resx</DependentUpon>";
  S = S + N + "    </Compile>";
  S = S + N + "    <Compile Include=\"Properties\\Settings.Designer.cs\">";
  S = S + N + "      <AutoGen>True</AutoGen>";
  S = S + N + "      <DependentUpon>Settings.settings</DependentUpon>";
  S = S + N + "      <DesignTimeSharedInput>True</DesignTimeSharedInput>";
  S = S + N + "    </Compile>";
  S = S + N + "    <EmbeddedResource Include=\"Properties\\Resources.resx\">";
  S = S + N + "      <Generator>ResXFileCodeGenerator</Generator>";
  S = S + N + "      <LastGenOutput>Resources.Designer.cs</LastGenOutput>";
  S = S + N + "    </EmbeddedResource>";
  S = S + N + "    <None Include=\"Properties\\Settings.settings\">";
  S = S + N + "      <Generator>SettingsSingleFileGenerator</Generator>";
  S = S + N + "      <LastGenOutput>Settings.Designer.cs</LastGenOutput>";
  S = S + N + "    </None>";
  S = S + N + "  </ItemGroup>";
  S = S + N + "  <ItemGroup>";
  S = S + N + "    <None Include=\"App.config\" />";
  S = S + N + "  </ItemGroup>";
  S = S + N + "  <ItemGroup>";
  S = S + N + "    <COMReference Include=\"ADODB\">";
  S = S + N + "      <Guid>{B691E011-1797-432E-907A-4D8C69339129}</Guid>";
  S = S + N + "      <VersionMajor>6</VersionMajor>";
  S = S + N + "      <VersionMinor>1</VersionMinor>";
  S = S + N + "      <Lcid>0</Lcid>";
  S = S + N + "      <WrapperTool>tlbimp</WrapperTool>";
  S = S + N + "      <Isolated>False</Isolated>";
  S = S + N + "      <EmbedInteropTypes>True</EmbedInteropTypes>";
  S = S + N + "    </COMReference>";
  S = S + N + "  </ItemGroup>";
  S = S + N + "  <Import Project=\"$(MSBuildToolsPath)\\Microsoft.CSharp.targets\" />";
  S = S + N + "</Project>";

  CreateProjectFile = S;

  WriteOut(ChgExt(FileName(vbpFile), ".csproj"), S);
  return CreateProjectFile;
}

public static string VBExtensionClass() {
  string VBExtensionClass = "";
  VBExtensionClass = ReadEntireFile(App.Path + "\\\\VBExtension.cs");
  return VBExtensionClass;
}

public static string VBAConstantsClass() {
  string VBAConstantsClass = "";
  VBAConstantsClass = ReadEntireFile(App.Path + "\\\\VBConstants.cs");
  return VBAConstantsClass;
}
}
