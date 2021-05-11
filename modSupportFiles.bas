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
  
  S = VBAConstantsClass()
  F = "VBConstants.cs"
  WriteOut F, S, ""
  
  S = AppConfigFile()
  F = "App.config"
  WriteOut F, S, ""

  S = AppXamlCsFile()
  F = "App.xaml.cs"
  WriteOut F, S, ""
  
  GeneratePropertiesFiles
End Function

Public Function GeneratePropertiesFiles()
  Dim S As String
  S = OutputFolder()
  S = S & "Properties\"
  If Dir(S, vbDirectory) = "" Then MkDir S
  
  
  WriteOut "Properties\Settings.settings", SettingsSettingsFile, "Properties"
  WriteOut "Properties\Settings.Designer.cs", SettingsDesignerCsFile, "Properties"
  WriteOut "Properties\AssemblyInfo.cs", AssemblyInfoFile, "Properties"
  WriteOut "Properties\Resources.resx", ResourcesResxFile, "Properties"
  WriteOut "Properties\Resources.Designer.cs", ResourcesDesignerCsFile, "Properties"

ResourcesResxFile
End Function

Public Function ApplicationXAML() As String
  Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "<Application x:Class=""Application"" "
  R = R & N & "xmlns = ""http://schemas.microsoft.com/winfx/2006/xaml/presentation"" "
  R = R & N & "xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"" "
  R = R & N & "xmlns:local=""clr-namespace:" & AssemblyName & """ "
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
  S = S & N & "    <RootNamespace>" & AssemblyName & "</RootNamespace>"
  S = S & N & "    <AssemblyName>" & AssemblyName & "</AssemblyName>"
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
  S = S & N & "    <Reference Include=""Microsoft.VisualBasic.PowerPacks, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL"" />"
  S = S & N & "    <Reference Include=""Microsoft.VisualBasic.PowerPacks.Vs, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL"" />"
  S = S & N & "    <Reference Include=""System"" />"
  S = S & N & "    <Reference Include=""System.Data"" />"
  S = S & N & "    <Reference Include=""System.Drawing"" />"
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
  S = S & N & "    <ApplicationDefinition Include=""Application.xaml"">"
  S = S & N & "      <Generator>MSBuild:Compile</Generator>"
  S = S & N & "      <SubType>Designer</SubType>"
  S = S & N & "    </ApplicationDefinition>"
  S = S & N & "    <Compile Include=""App.xaml.cs"">"
  S = S & N & "      <DependentUpon>App.xaml</DependentUpon>"
  S = S & N & "      <SubType>Code</SubType>"
  S = S & N & "    </Compile>"
  
  For Each L In Split(VBPForms(vbpFile), vbCrLf)
  If L = "" Then GoTo SkipForm
  S = S & N & "    <Page Include=""" & OutputSubFolder(L) & ChgExt(L, ".xaml") & """>"
  S = S & N & "      <SubType>Designer</SubType>"
  S = S & N & "      <Generator>MSBuild:Compile</Generator>"
  S = S & N & "    </Page>"
  S = S & N & "    <Compile Include=""" & OutputSubFolder(L) & ChgExt(L, ".xaml.cs") & """>"
  S = S & N & "      <DependentUpon>" & ChgExt(L, ".xaml") & "</DependentUpon>"
  S = S & N & "      <SubType>Code</SubType>"
  S = S & N & "    </Compile>"
SkipForm:
  Next

  
  S = S & N & "    <Compile Include=""VBExtension.cs"" />"
  S = S & N & "    <Compile Include=""VBConstants.cs"" />"
  For Each L In Split(VBPClasses(vbpFile) & vbCrLf & VBPModules(vbpFile), vbCrLf)
If L = "" Then GoTo SkipClass
  S = S & N & "    <Compile Include=""" & OutputSubFolder(L) & ChgExt(L, ".cs") & """ />"
SkipClass:
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
  S = S & N & "  <ItemGroup>"
  S = S & N & "    <COMReference Include=""ADODB"">"
  S = S & N & "      <Guid>{B691E011-1797-432E-907A-4D8C69339129}</Guid>"
  S = S & N & "      <VersionMajor>6</VersionMajor>"
  S = S & N & "      <VersionMinor>1</VersionMinor>"
  S = S & N & "      <Lcid>0</Lcid>"
  S = S & N & "      <WrapperTool>tlbimp</WrapperTool>"
  S = S & N & "      <Isolated>False</Isolated>"
  S = S & N & "      <EmbedInteropTypes>True</EmbedInteropTypes>"
  S = S & N & "    </COMReference>"
  S = S & N & "  </ItemGroup>"
  S = S & N & "  <Import Project=""$(MSBuildToolsPath)\Microsoft.CSharp.targets"" />"
  S = S & N & "</Project>"
  
  CreateProjectFile = S
  
  WriteOut ChgExt(tFileName(vbpFile), ".csproj"), S
End Function

Public Function VBExtensionClass() As String
  VBExtensionClass = ReadEntireFile(App.Path & "\\VBExtension.cs")
End Function

Public Function VBAConstantsClass() As String
  VBAConstantsClass = ReadEntireFile(App.Path & "\\VBConstants.cs")
End Function

Public Function AppConfigFile() As String
  Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "<?xml version='1.0' encoding='utf-8'?>"
  R = R & N & "<configuration>"
  R = R & N & "    <startup>"
  R = R & N & "        <supportedRuntime version='v4.0' sku='.NETFramework,Version=v4.7.2'/>"
  R = R & N & "    </startup>"
  R = R & N & "</configuration>"

  AppConfigFile = R
End Function

Public Function AppXamlCsFile() As String
 Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "using System;"
  R = R & N & "using System.Collections.Generic;"
  R = R & N & "using System.Configuration;"
  R = R & N & "using System.Data;"
  R = R & N & "using System.Linq;"
  R = R & N & "using System.Threading.Tasks;"
  R = R & N & "using System.Windows;"
  R = R & N & ""
  R = R & N & "namespace " & AssemblyName
  R = R & N & "{"
  R = R & N & "  /// <summary>"
  R = R & N & "  /// Interaction logic for App.xaml"
  R = R & N & "  /// </summary>"
  R = R & N & "  public partial class App : Application"
  R = R & N & "    {"
  R = R & N & "    }"
  R = R & N & "}"
  R = R & N & ""

  AppXamlCsFile = R
End Function

Public Function SettingsSettingsFile() As String
  Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "<?xml version='1.0' encoding='utf-8'?>"
  R = R & N & "<SettingsFile xmlns='uri:settings' CurrentProfile='(Default)'>"
  R = R & N & "  <Profiles>"
  R = R & N & "    <Profile Name='(Default)' />"
  R = R & N & "  </Profiles>"
  R = R & N & "  <Settings />"
  R = R & N & "</SettingsFile>"

  SettingsSettingsFile = R
End Function

Public Function SettingsDesignerCsFile() As String
  Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "//------------------------------------------------------------------------------"
  R = R & N & "// <auto-generated>"
  R = R & N & "//     This code was generated by a tool."
  R = R & N & "//     Runtime Version:4.0.30319.42000"
  R = R & N & "//"
  R = R & N & "//     Changes to this file may cause incorrect behavior and will be lost if"
  R = R & N & "//     the code is regenerated."
  R = R & N & "// </auto-generated>"
  R = R & N & "//------------------------------------------------------------------------------"
  R = R & N & ""
  R = R & N & "namespace " & AssemblyName & ".Properties {"
  R = R & N & ""
  R = R & N & ""
  R = R & N & "    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]"
  R = R & N & "    [global::System.CodeDom.Compiler.GeneratedCodeAttribute(""Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator"", ""15.9.0.0"")]"
  R = R & N & "    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {"
  R = R & N & ""
  R = R & N & "        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));"
  R = R & N & ""
  R = R & N & "        public static Settings Default {"
  R = R & N & "            get {"
  R = R & N & "                return defaultInstance;"
  R = R & N & "            }"
  R = R & N & "        }"
  R = R & N & "    }"
  R = R & N & "}"

  SettingsDesignerCsFile = R
End Function

Public Function AssemblyInfoFile()
  Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  
  R = R & M & "using System.Reflection;"
  R = R & N & "using System.Resources;"
  R = R & N & "using System.Runtime.CompilerServices;"
  R = R & N & "using System.Runtime.InteropServices;"
  R = R & N & "using System.Windows;"
  R = R & N & ""
  R = R & N & "// General Information about an assembly is controlled through the following"
  R = R & N & "// set of attributes. Change these attribute values to modify the information"
  R = R & N & "// associated with an assembly."
  R = R & N & "[assembly: AssemblyTitle(""" & AssemblyName & """)]"
  R = R & N & "[assembly: AssemblyDescription("""")]"
  R = R & N & "[assembly: AssemblyConfiguration("""")]"
  R = R & N & "[assembly: AssemblyCompany("""")]"
  R = R & N & "[assembly: AssemblyProduct(""" & AssemblyName & """)]"
  R = R & N & "[assembly: AssemblyCopyright(""Copyright " & Year(Now) & """)]"
  R = R & N & "[assembly: AssemblyTrademark("""")]"
  R = R & N & "[assembly: AssemblyCulture("""")]"
  R = R & N & ""
  R = R & N & "// Setting ComVisible to false makes the types in this assembly not visible"
  R = R & N & "// to COM components.  If you need to access a type in this assembly from"
  R = R & N & "// COM, set the ComVisible attribute to true on that type."
  R = R & N & "[assembly: ComVisible(false)]"
  R = R & N & ""
  R = R & N & "//In order to begin building localizable applications, set"
  R = R & N & "//<UICulture>CultureYouAreCodingWith</UICulture> in your .csproj file"
  R = R & N & "//inside a <PropertyGroup>.  For example, if you are using US english"
  R = R & N & "//in your source files, set the <UICulture> to en-US.  Then uncomment"
  R = R & N & "//the NeutralResourceLanguage attribute below.  Update the ""en-US"" in"
  R = R & N & "//the line below to match the UICulture setting in the project file."
  R = R & N & ""
  R = R & N & "//[assembly: NeutralResourcesLanguage(""en-US"", UltimateResourceFallbackLocation.Satellite)]"
  R = R & N & ""
  R = R & N & ""
  R = R & N & "[assembly: ThemeInfo("
  R = R & N & "  ResourceDictionaryLocation.None, //where theme specific resource dictionaries are located"
  R = R & N & "//(used if a resource is not found in the page,"
  R = R & N & "// or application resource dictionaries)"
  R = R & N & "ResourceDictionaryLocation.SourceAssembly //where the generic resource dictionary is located"
  R = R & N & "                                              //(used if a resource is not found in the page,"
  R = R & N & "                                              // app, or any theme specific resource dictionaries)"
  R = R & N & ")]"
  R = R & N & ""
  R = R & N & ""
  R = R & N & "// Version information for an assembly consists of the following four values:"
  R = R & N & "//"
  R = R & N & "//      Major Version"
  R = R & N & "//      Minor Version"
  R = R & N & "//      Build Number"
  R = R & N & "//      Revision"
  R = R & N & "//"
  R = R & N & "// You can specify all the values or you can default the Build and Revision Numbers"
  R = R & N & "// by using the '*' as shown below:"
  R = R & N & "// [assembly: AssemblyVersion(""1.0.*"")]"
  R = R & N & "[assembly: AssemblyVersion(""1.0.0.0"")]"
  R = R & N & "[assembly: AssemblyFileVersion(""1.0.0.0"")]"
  
  AssemblyInfoFile = R
End Function

Public Function ResourcesResxFile()
  Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  

  R = R & N & "<?xml version='1.0' encoding='utf-8'?>"
  R = R & N & "<root>"
  R = R & N & "  <!--"
  R = R & N & "    Microsoft ResX Schema"
  R = R & N & ""
  R = R & N & "    Version 2.0"
  R = R & N & ""
  R = R & N & "    The primary goals of this format is to allow a simple XML format"
  R = R & N & "    that is mostly human readable. The generation and parsing of the"
  R = R & N & "    various data types are done through the TypeConverter classes"
  R = R & N & "    associated with the data types."
  R = R & N & ""
  R = R & N & "    Example:"
  R = R & N & ""
  R = R & N & "    ... ado.net/XML headers & schema ..."
  R = R & N & "    <resheader name='resmimetype'>text/microsoft-resx</resheader>"
  R = R & N & "    <resheader name='version'>2.0</resheader>"
  R = R & N & "    <resheader name='reader'>System.Resources.ResXResourceReader, System.Windows.Forms, ...</resheader>"
  R = R & N & "    <resheader name='writer'>System.Resources.ResXResourceWriter, System.Windows.Forms, ...</resheader>"
  R = R & N & "    <data name='Name1'><value>this is my long string</value><comment>this is a comment</comment></data>"
  R = R & N & "    <data name='Color1' type='System.Drawing.Color, System.Drawing'>Blue</data>"
  R = R & N & "    <data name='Bitmap1' mimetype='application/x-microsoft.net.object.binary.base64'>"
  R = R & N & "        <value>[base64 mime encoded serialized .NET Framework object]</value>"
  R = R & N & "    </data>"
  R = R & N & "    <data name='Icon1' type='System.Drawing.Icon, System.Drawing' mimetype='application/x-microsoft.net.object.bytearray.base64'>"
  R = R & N & "        <value>[base64 mime encoded string representing a byte array form of the .NET Framework object]</value>"
  R = R & N & "        <comment>This is a comment</comment>"
  R = R & N & "    </data>"
  R = R & N & ""
  R = R & N & "    There are any number of 'resheader' rows that contain simple"
  R = R & N & "    name/value pairs."
  R = R & N & ""
  R = R & N & "    Each data row contains a name, and value. The row also contains a"
  R = R & N & "    type or mimetype. Type corresponds to a .NET class that support"
  R = R & N & "    text/value conversion through the TypeConverter architecture."
  R = R & N & "    Classes that don't support this are serialized and stored with the"
  R = R & N & "    mimetype set."
  R = R & N & ""
  R = R & N & "    The mimetype is used for serialized objects, and tells the"
  R = R & N & "    ResXResourceReader how to depersist the object. This is currently not"
  R = R & N & "    extensible. For a given mimetype the value must be set accordingly:"
  R = R & N & ""
  R = R & N & "    Note - application/x-microsoft.net.object.binary.base64 is the format"
  R = R & N & "    that the ResXResourceWriter will generate, however the reader can"
  R = R & N & "    read any of the formats listed below."
  R = R & N & ""
  R = R & N & "    mimetype: application/x-microsoft.net.object.binary.base64"
  R = R & N & "    value   : The object must be serialized with"
  R = R & N & "            : System.Serialization.Formatters.Binary.BinaryFormatter"
  R = R & N & "            : and then encoded with base64 encoding."
  R = R & N & ""
  R = R & N & "    mimetype: application/x-microsoft.net.object.soap.base64"
  R = R & N & "    value   : The object must be serialized with"
  R = R & N & "            : System.Runtime.Serialization.Formatters.Soap.SoapFormatter"
  R = R & N & "            : and then encoded with base64 encoding."
  R = R & N & ""
  R = R & N & "    mimetype: application/x-microsoft.net.object.bytearray.base64"
  R = R & N & "    value   : The object must be serialized into a byte array"
  R = R & N & "            : using a System.ComponentModel.TypeConverter"
  R = R & N & "            : and then encoded with base64 encoding."
  R = R & N & "    -->"
  R = R & N & "  <xsd:schema id='root' xmlns='' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:msdata='urn:schemas-microsoft-com:xml-msdata'>"
  R = R & N & "    <xsd:element name='root' msdata:IsDataSet='true'>"
  R = R & N & "      <xsd:complexType>"
  R = R & N & "        <xsd:choice maxOccurs='unbounded'>"
  R = R & N & "          <xsd:element name='metadata'>"
  R = R & N & "            <xsd:complexType>"
  R = R & N & "              <xsd:sequence>"
  R = R & N & "                <xsd:element name='value' type='xsd:string' minOccurs='0' />"
  R = R & N & "              </xsd:sequence>"
  R = R & N & "              <xsd:attribute name='name' type='xsd:string' />"
  R = R & N & "              <xsd:attribute name='type' type='xsd:string' />"
  R = R & N & "              <xsd:attribute name='mimetype' type='xsd:string' />"
  R = R & N & "            </xsd:complexType>"
  R = R & N & "          </xsd:element>"
  R = R & N & "          <xsd:element name='assembly'>"
  R = R & N & "            <xsd:complexType>"
  R = R & N & "              <xsd:attribute name='alias' type='xsd:string' />"
  R = R & N & "              <xsd:attribute name='name' type='xsd:string' />"
  R = R & N & "            </xsd:complexType>"
  R = R & N & "          </xsd:element>"
  R = R & N & "          <xsd:element name='data'>"
  R = R & N & "            <xsd:complexType>"
  R = R & N & "              <xsd:sequence>"
  R = R & N & "                <xsd:element name='value' type='xsd:string' minOccurs='0' msdata:Ordinal='1' />"
  R = R & N & "                <xsd:element name='comment' type='xsd:string' minOccurs='0' msdata:Ordinal='2' />"
  R = R & N & "              </xsd:sequence>"
  R = R & N & "              <xsd:attribute name='name' type='xsd:string' msdata:Ordinal='1' />"
  R = R & N & "              <xsd:attribute name='type' type='xsd:string' msdata:Ordinal='3' />"
  R = R & N & "              <xsd:attribute name='mimetype' type='xsd:string' msdata:Ordinal='4' />"
  R = R & N & "            </xsd:complexType>"
  R = R & N & "          </xsd:element>"
  R = R & N & "          <xsd:element name='resheader'>"
  R = R & N & "            <xsd:complexType>"
  R = R & N & "              <xsd:sequence>"
  R = R & N & "                <xsd:element name='value' type='xsd:string' minOccurs='0' msdata:Ordinal='1' />"
  R = R & N & "              </xsd:sequence>"
  R = R & N & "              <xsd:attribute name='name' type='xsd:string' use='required' />"
  R = R & N & "            </xsd:complexType>"
  R = R & N & "          </xsd:element>"
  R = R & N & "        </xsd:choice>"
  R = R & N & "      </xsd:complexType>"
  R = R & N & "    </xsd:element>"
  R = R & N & "  </xsd:schema>"
  R = R & N & "  <resheader name='resmimetype'>"
  R = R & N & "    <value>text/microsoft-resx</value>"
  R = R & N & "  </resheader>"
  R = R & N & "  <resheader name='version'>"
  R = R & N & "    <value>2.0</value>"
  R = R & N & "  </resheader>"
  R = R & N & "  <resheader name='reader'>"
  R = R & N & "    <value>System.Resources.ResXResourceReader, System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>"
  R = R & N & "  </resheader>"
  R = R & N & "  <resheader name='writer'>"
  R = R & N & "    <value>System.Resources.ResXResourceWriter, System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>"
  R = R & N & "  </resheader>"
  R = R & N & "</root>End Function"
  
  ResourcesResxFile = R
End Function

Public Function ResourcesDesignerCsFile()
  Dim R As String, M As String, N As String
  R = "": M = "": N = vbCrLf
  
  R = R & N & "//------------------------------------------------------------------------------"
  R = R & N & "// <auto-generated>"
  R = R & N & "//     This code was generated by a tool."
  R = R & N & "//     Runtime Version:4.0.30319.42000"
  R = R & N & "//"
  R = R & N & "//     Changes to this file may cause incorrect behavior and will be lost if"
  R = R & N & "//     the code is regenerated."
  R = R & N & "// </auto-generated>"
  R = R & N & "//------------------------------------------------------------------------------"
  R = R & N & ""
  R = R & N & "namespace " & AssemblyName & ".Properties {"
  R = R & N & "    using System;"
  R = R & N & ""
  R = R & N & ""
  R = R & N & "    /// <summary>"
  R = R & N & "    ///   A strongly-typed resource class, for looking up localized strings, etc."
  R = R & N & "    /// </summary>"
  R = R & N & "    // This class was auto-generated by the StronglyTypedResourceBuilder"
  R = R & N & "    // class via a tool like ResGen or Visual Studio."
  R = R & N & "    // To add or remove a member, edit your .ResX file then rerun ResGen"
  R = R & N & "    // with the /str option, or rebuild your VS project."
  R = R & N & "    [global::System.CodeDom.Compiler.GeneratedCodeAttribute(""System.Resources.Tools.StronglyTypedResourceBuilder"", ""15.0.0.0"")]"
  R = R & N & "    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]"
  R = R & N & "    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]"
  R = R & N & "    internal class Resources {"
  R = R & N & ""
  R = R & N & "        private static global::System.Resources.ResourceManager resourceMan;"
  R = R & N & ""
  R = R & N & "        private static global::System.Globalization.CultureInfo resourceCulture;"
  R = R & N & ""
  R = R & N & "        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(""Microsoft.Performance"", ""CA1811:AvoidUncalledPrivateCode"")]"
  R = R & N & "        internal Resources() {"
  R = R & N & "        }"
  R = R & N & ""
  R = R & N & "        /// <summary>"
  R = R & N & "        ///   Returns the cached ResourceManager instance used by this class."
  R = R & N & "        /// </summary>"
  R = R & N & "        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]"
  R = R & N & "        internal static global::System.Resources.ResourceManager ResourceManager {"
  R = R & N & "            get {"
  R = R & N & "                if (object.ReferenceEquals(resourceMan, null)) {"
  R = R & N & "                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager(""WinCDS.Properties.Resources"", typeof(Resources).Assembly);"
  R = R & N & "                    resourceMan = temp;"
  R = R & N & "                }"
  R = R & N & "                return resourceMan;"
  R = R & N & "            }"
  R = R & N & "        }"
  R = R & N & ""
  R = R & N & "        /// <summary>"
  R = R & N & "        ///   Overrides the current thread's CurrentUICulture property for all"
  R = R & N & "        ///   resource lookups using this strongly typed resource class."
  R = R & N & "        /// </summary>"
  R = R & N & "        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]"
  R = R & N & "        internal static global::System.Globalization.CultureInfo Culture {"
  R = R & N & "            get {"
  R = R & N & "                return resourceCulture;"
  R = R & N & "            }"
  R = R & N & "            set {"
  R = R & N & "                resourceCulture = value;"
  R = R & N & "            }"
  R = R & N & "        }"
  R = R & N & "    }"
  R = R & N & "}"
  
  ResourcesDesignerCsFile = R
End Function
