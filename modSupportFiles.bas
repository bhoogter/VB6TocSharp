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
  S = S & N & "    <Compile Include=""VBConstants.cs"" />"
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
  S = S & N & "  public enum vbTriState {  vbFalse = 0, vbTrue = -1, vbUseDefault = -2 }"
  S = S & N
  S = S & N & "  public static object Printer;"
  S = S & N
  S = S & N & "  public static void Unload(dynamic Ob) { Ob.Close(); }"
  S = S & N
  S = S & N & "  public static object IIf(bool A, object B, object C) { return !!A ? B : C; }"
  S = S & N & "  public static bool IIf(bool A, bool B, bool C) { return !!A ? B : C; }"
  S = S & N & "  public static string IIf(bool A, string B, string C) { return !!A ? B : C; }"
  S = S & N & "  public static double IIf(bool A, double B, double C) { return !!A ? B : C; }"
  S = S & N & "  public static decimal IIf(bool A, decimal B, decimal C) { return !!A ? B : C; }"
  S = S & N & "  public static long IIf(bool A, long B, long C) { return !!A ? B : C; }"
  S = S & N & "  public static DateTime IIf(bool A, DateTime B, DateTime C) { return !!A ? B : C; }"
  S = S & N
  S = S & N & "  public static bool IsMissing(object A) { return false; }"
  S = S & N & "  public static bool IsNull(object A) { return A == null || (A is System.DBNull); }"
  S = S & N & "  public static bool IsNothing(object A) { return IsNull(A); }"
  S = S & N & "  public static bool IsObject(object A) { return !IsNothing(A); }"
  S = S & N
  S = S & N & "  public static System.DateTime NullDate() { try { return System.DateTime.Parse(""1/1/2001""); } catch { return DateTime.Today; } }"
  S = S & N & "  public static bool IsDate(string D) { try { System.DateTime.Parse(D); } catch { return false; } return true; }"
  S = S & N
  S = S & N & "  public static System.DateTime CDate(object A) { return IsDate(A.ToString()) ? NullDate() : System.DateTime.Parse(A.ToString()); }"
  S = S & N & "  public static double CDbl(object A)  { return (A is System.IConvertible) ? ((System.IConvertible)A).ToDouble(null) : 0; }"
  S = S & N & "  public static double CLng(object A) { return (A is System.IConvertible) ? ((System.IConvertible)A).ToInt64(null) : 0; }"
  S = S & N & "  public static double CInt(object A) { return (A is System.IConvertible) ? ((System.IConvertible)A).ToInt32(null) : 0; }"
  S = S & N & "  public static string CStr(object A) { return A.ToString(); }"
  S = S & N & "  public static bool CBool(object A) { { return (A is System.IConvertible) ? ((System.IConvertible)A).ToBoolean(null) : false; } }"
  S = S & N
  S = S & N & "  public static System.DateTime DateValue(object A) { return CDate(A); }"
  S = S & N
  S = S & N & "  public static bool IsList(object A) { return A != null && (A is System.Collections.IList); }"
  S = S & N & "  public static long LBound(object A) { return 0; }"
  S = S & N & "  public static long UBound(object A) { return A != null && (A is System.Collections.IList) ? ((System.Collections.IList)A).Count - 1 : 0; }"
  S = S & N
  S = S & N & "  public static bool IsLike(string A, string B) { return Microsoft.VisualBasic.CompilerServices.LikeOperator.LikeString(A, B, Microsoft.VisualBasic.CompareMethod.Binary); }"
  S = S & N & "  public static dynamic Switch(params dynamic[] list) { for (long i = 0; i < list.Length; i += 2) if ((bool)list[i * 2]) return list[i * 2 + 1]; return null; }"
  S = S & N & "  public static dynamic Choose(long Idx, params dynamic[] list) { if (Idx < 0 || Idx >= list.Length) return null; return list[Idx]; }"
  S = S & N
  S = S & N & "  public static bool VBOpenFile(dynamic A, dynamic B) { return false; }"
  S = S & N & "  public static bool VBWriteFile(dynamic A, dynamic B) { return false; }"
  S = S & N & "  public static bool VBCloseFile(dynamic A) { return false; }"
  S = S & N & "  public static string VBReadFileLine(dynamic A, dynamic B) { return """"; }"
  S = S & N & "  public static bool DoEvents() { return false; }"
  S = S & N
  S = S & N & "  public static bool Resume() { return false; }"
  S = S & N & "  public static bool End() { return false; }"
  S = S & M
  S = S & N & "  public static bool HasEmptyText(this TextBox textBox) { return string.IsNullOrEmpty(textBox.Text); }"
  S = S & N & "  public static decimal getValue(this TextBox textBox) { try { return Decimal.Parse(textBox.Text); } catch { return 0; } }"
  S = S & N & "  public static decimal setValue(this TextBox textBox, decimal value) { textBox.Text = value.ToString(); return getValue(textBox); }"
  S = S & N & "  public static long getValueLong(this TextBox textBox) { try { return long.Parse(textBox.Text); } catch { return 0; } }"
  S = S & N & "  public static long setValueLong(this TextBox textBox, long value) { textBox.Text = value.ToString(); return getValueLong(textBox); }"
  
  S = S & N & ""
  S = S & N & "  public static bool getValue(this CheckBox chk) { try { return ((bool)chk.IsChecked); } catch { return false; } }"
  S = S & N & "  public static bool setValue(this CheckBox chk, bool value) { chk.IsChecked = value; return getValue(chk); }"
  S = S & N & "//    public static long getValue(this CheckBox chk) { try { return ((bool)chk.IsChecked); } catch { return false; } }"
  S = S & N & "    public static long setValue(this CheckBox chk, long value) { chk.IsChecked = value != 1; return getValue(chk) ? 1: 0; }"

  S = S & N & ""
  S = S & N & "    public static bool getValue(this Button btn) { try { return ((bool) btn.IsPressed); } catch { return false; } }"
  S = S & N & "    public static bool setValue(this Button btn, bool value) { try { btn.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); return true; } catch { return false; } }"

  S = S & N
  S = S & N & "  public static bool getVisible(this Control c) { return c.Visibility == System.Windows.Visibility.Visible; }"
  S = S & N & "  public static bool setVisible(this Control c, bool value) { c.Visibility = value ? System.Windows.Visibility.Visible : System.Windows.Visibility.Hidden; return c.getVisible(); }"
  S = S & N & "  public static bool SetFocus(this Control c) { try { return c.Focus(); } catch { return false; } }"
  S = S & N & "  public static bool Move(this Control c, decimal X = -10000, decimal Y = -10000, decimal W = -1000, decimal H = -10000, bool MakeVisible = false ) {"
  S = S & N & "      if (W != -10000) c.Height = (double) W;"
  S = S & N & "      if (H != -10000) c.Height = (double) H;"
  S = S & N & "      c.Margin = new System.Windows.Thickness(X == -10000 ? c.Margin.Left : (double)X, Y == -10000 ? c.Margin.Top : (double)Y, c.Width, c.Height);"
  S = S & N & "      try { return c.Focus(); } catch { return false; }"
  S = S & N & "  }"
  
  S = S & N
  S = S & N & "}"
  
  VBExtensionClass = S
End Function

Public Function VBAConstantsClass() As String
  Dim S As String, M As String, N As String
  Dim L
  S = ""
  M = ""
  N = vbCrLf
  
  S = S & M & ""
  S = S & M & UsingEverything
  S = S & N
  S = S & M & "public static class VBConstants {"
  S = S & N & "  public const long vbKeyLButton = 1; // Left mouse button"

  S = S & N & "    public const long vbKeyRButton = 2;  // CANCEL mouse button "
  S = S & N & "    public const long vbKeyCancel = 3;  // Middle key  "
  S = S & N & "    public const long vbKeyMButton = 4;  // BACKSPACE mouse button "
  S = S & N & "    public const long vbKeyBack = 8;  // TAB key  "
  S = S & N & "    public const long vbKeyTab = 9;  //  key  "
  S = S & N & "    public const long vbKeyClear = 12;  //  CLEAR key "
  S = S & N & "    public const long vbKeyReturn = 13;  //  ENTER key "
  S = S & N & "    public const long vbKeyShift = 16;  //  SHIFT key "
  S = S & N & "    public const long vbKeyControl = 17;  //  CTRL key "
  S = S & N & "    public const long vbKeyMenu = 18;  //  MENU key "
  S = S & N & "    public const long vbKeyPause = 19;  //  PAUSE key "
  S = S & N & "    public const long vbKeyCapital = 20;  //  CAPS lock key"
  S = S & N & "    public const long vbKeyEscape = 27;  //  ESC key "
  S = S & N & "    public const long vbKeySpace = 32;  //  SPACEBAR key "
  S = S & N & "    public const long vbKeyPageUp = 33;  //  PAGE UP key"
  S = S & N & "    public const long vbKeyPageDown = 34;  //  PAGE DOWN key"
  S = S & N & "    public const long vbKeyEnd = 35;  //  END key "
  S = S & N & "    public const long vbKeyHome = 36;  //  HOME key "
  S = S & N & "    public const long vbKeyLeft = 37;  //  LEFT ARROW key"
  S = S & N & "    public const long vbKeyUp = 38;  //  UP ARROW key"
  S = S & N & "    public const long vbKeyRight = 39;  //  RIGHT ARROW key"
  S = S & N & "    public const long vbKeyDown = 40;  //  DOWN ARROW key"
  S = S & N & "    public const long vbKeySelect = 41;  //  SELECT key "
  S = S & N & "    public const long vbKeyPrint = 42;  //  print SCREEN key"
  S = S & N & "    public const long vbKeyExecute = 43;  //  EXECUTE key "
  S = S & N & "    public const long vbKeySnapshot = 44;  //  SNAPSHOT key "
  S = S & N & "    public const long vbKeyInsert = 45;  //  INS key "
  S = S & N & "    public const long vbKeyDelete = 46;  //  DEL key "
  S = S & N & "    public const long vbKeyHelp = 47;  // NUM HELP key "
  S = S & N & "    public const long vbKeyNumlock = 144;  //  lock key "
  S = S & N & "    public const long vbKeyA = 65;  //  A key "
  S = S & N & "    public const long vbKeyB = 66;  //  B key "
  S = S & N & "    public const long vbKeyC = 67;  //  C key "
  S = S & N & "    public const long vbKeyD = 68;  //  D key "
  S = S & N & "    public const long vbKeyE = 69;  //  E key "
  S = S & N & "    public const long vbKeyF = 70;  //  F key "
  S = S & N & "    public const long vbKeyG = 71;  //  G key "
  S = S & N & "    public const long vbKeyH = 72;  //  H key "
  S = S & N & "    public const long vbKeyI = 73;  //  I key "
  S = S & N & "    public const long vbKeyJ = 74;  //  J key "
  S = S & N & "    public const long vbKeyK = 75;  //  K key "
  S = S & N & "    public const long vbKeyL = 76;  //  L key "
  S = S & N & "    public const long vbKeyM = 77;  //  M key "
  S = S & N & "    public const long vbKeyN = 78;  //  N key "
  S = S & N & "    public const long vbKeyO = 79;  //  O key "
  S = S & N & "    public const long vbKeyP = 80;  //  P key "
  S = S & N & "    public const long vbKeyQ = 81;  //  Q key "
  S = S & N & "    public const long vbKeyR = 82;  //  R key "
  S = S & N & "    public const long vbKeyS = 83;  //  S key "
  S = S & N & "    public const long vbKeyT = 84;  //  T key "
  S = S & N & "    public const long vbKeyU = 85;  //  U key "
  S = S & N & "    public const long vbKeyV = 86;  //  V key "
  S = S & N & "    public const long vbKeyW = 87;  //  W key "
  S = S & N & "    public const long vbKeyX = 88;  //  X key "
  S = S & N & "    public const long vbKeyY = 89;  //  Y key "
  S = S & N & "    public const long vbKeyZ = 90;  //  Z key "
  S = S & N & "    public const long vbKey0 = 48;  //  0 key "
  S = S & N & "    public const long vbKey1 = 49;  //  1 key "
  S = S & N & "    public const long vbKey2 = 50;  //  2 key "
  S = S & N & "    public const long vbKey3 = 51;  //  3 key "
  S = S & N & "    public const long vbKey4 = 52;  //  4 key "
  S = S & N & "    public const long vbKey5 = 53;  //  5 key "
  S = S & N & "    public const long vbKey6 = 54;  //  6 key "
  S = S & N & "    public const long vbKey7 = 55;  //  7 key "
  S = S & N & "    public const long vbKey8 = 56;  //  8 key "
  S = S & N & "    public const long vbKey9 = 57;  //  9 key "
  S = S & N & "    public const long vbKeyNumpad0 = 96;  //  0 key "
  S = S & N & "    public const long vbKeyNumpad1 = 97;  //  1 key "
  S = S & N & "    public const long vbKeyNumpad2 = 98;  //  2 key "
  S = S & N & "    public const long vbKeyNumpad3 = 99;  // 4 3 key "
  S = S & N & "    public const long vbKeyNumpad4 = 100;  // 5 key  "
  S = S & N & "    public const long vbKeyNumpad5 = 101;  // 6 key  "
  S = S & N & "    public const long vbKeyNumpad6 = 102;  // 7 key  "
  S = S & N & "    public const long vbKeyNumpad7 = 103;  // 8 key  "
  S = S & N & "    public const long vbKeyNumpad8 = 104;  // 9 key  "
  S = S & N & "    public const long vbKeyNumpad9 = 105;  // MULTIPLICATION key  "
  S = S & N & "    public const long vbKeyMultiply = 106;  // PLUS SIGN (*) key"
  S = S & N & "    public const long vbKeyAdd = 107;  // ENTER SIGN (+) key"
  S = S & N & "    public const long vbKeySeparator = 108;  // MINUS (keypad) key "
  S = S & N & "    public const long vbKeySubtract = 109;  // DECIMAL SIGN (-) key"
  S = S & N & "    public const long vbKeyDecimal = 110;  // DIVISION POINT(.) key "
  S = S & N & "    public const long vbKeyDivide = 111;  // F1 SIGN (/) key"
  S = S & N & "    public const long vbKeyF1 = 112;  // F2 key  "
  S = S & N & "    public const long vbKeyF2 = 113;  // F3 key  "
  S = S & N & "    public const long vbKeyF3 = 114;  // F4 key  "
  S = S & N & "    public const long vbKeyF4 = 115;  // F5 key  "
  S = S & N & "    public const long vbKeyF5 = 116;  // F6 key  "
  S = S & N & "    public const long vbKeyF6 = 117;  // F7 key  "
  S = S & N & "    public const long vbKeyF7 = 118;  // F8 key  "
  S = S & N & "    public const long vbKeyF8 = 119;  // F9 key  "
  S = S & N & "    public const long vbKeyF9 = 120;  // F10 key  "
  S = S & N & "    public const long vbKeyF10 = 121;  // F11 key  "
  S = S & N & "    public const long vbKeyF11 = 122;  // F12 key  "
  S = S & N & "    public const long vbKeyF12 = 123;  // F13 key  "
  S = S & N & "    public const long vbKeyF13 = 124;  // F14 key  "
  S = S & N & "    public const long vbKeyF14 = 125;  // F15 key  "
  S = S & N & "    public const long vbKeyF15 = 126;  // F16 key  "
  S = S & N & "    public const long vbKeyF16 = 127;  //  key  "
  
  S = S & N & "    public const long vbBlack = 0x0;  // BLACK"
  S = S & N & "    public const long vbBlue = 0x0000FF;  // BLUE"
  S = S & N & "    public const long vbCyan = 0x00FFFF;  // CYAN"
  S = S & N & "    public const long vbGreen = 0x00FF00;  // GREEN"
  S = S & N & "    public const long vbMagenta = 0xFFFF00;  // MAGENTA"
  S = S & N & "    public const long vbRed = 0xFF0000;  // RED"
  S = S & N & "    public const long vbWhite = 0xFFFFFF;  // WHITE"
  S = S & N & "    public const long vbYellow = 0xFF00FF;  // YELLOW"
  
  S = S & N & "    public const long vbModal = 0x1;"
  
  S = S & N & "    public const long vbAlignNone = 0;"
  S = S & N & "    public const long vbAlignTop = 1;"
  S = S & N & "    public const long vbAlignBottom = 2;"
  S = S & N & "    public const long vbAlignLeft = 3;"
  S = S & N & "    public const long vbAlignRight = 4;"
  
  
  S = S & N & " }"
  
  VBAConstantsClass = S
End Function

