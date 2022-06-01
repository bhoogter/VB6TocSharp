using System;
using System.Collections.Generic;
using VB2CS.Forms;
using static Microsoft.VisualBasic.FileSystem;
using static modGit;
using static VBExtension;



static class modTestCases
{
    // @NO-LINT
    // This module exists solely to list test conversion caess to make sure the converter can convert itself containing them.
    // There should be no active and/or used code in this module.
    // These tests are not run, they are conversion tests.  They should be converted correctly when this project is converted.
    public static void testCallModuleFunction()
    {
        // module name (w/, w/o)
        // assign value (w/w/o)
        // empty args parans (w/w/o)
        string S = "";
        modGit.GitVersion();
        GitVersion();
        modGit.GitCmd("git --version");
        S = modGit.GitCmd("git --verison");
        GitCmd("git --version");
        S = GitCmd("git --verison");
        S = modGit.GitVersion();
        modGit.GitVersion();
        S = GitVersion();
        GitVersion();
    }
    public static void testBooleans()
    {
        // not (w/w/o)
        // if (w/w/o)
        // fcall (w/w/o)
        bool B = false;
        bool NB = false;
        B = HasGit();
        B = HasGit();
        B = modGit.HasGit();
        B = modGit.HasGit();
        B = !HasGit();
        B = !HasGit();
        B = !modGit.HasGit();
        B = !modGit.HasGit();
        TestCallWithBooleanFunction(HasGit());
        TestCallWithBooleanFunction(!HasGit());
        TestCallWithBooleanFunction(modGit.HasGit());
        TestCallWithBooleanFunction(!modGit.HasGit());
        TestCallWithBooleanFunction(HasGit());
        TestCallWithBooleanFunction(!HasGit());
        TestCallWithBooleanFunction(modGit.HasGit());
        TestCallWithBooleanFunction(!modGit.HasGit());
        if (HasGit()) Console.WriteLine("");
        if (HasGit()) Console.WriteLine("");
        if (modGit.HasGit()) Console.WriteLine("");
        if (modGit.HasGit()) Console.WriteLine("");
        if (!HasGit()) Console.WriteLine("");
        if (!HasGit()) Console.WriteLine("");
        if (!modGit.HasGit()) Console.WriteLine("");
        if (!modGit.HasGit()) Console.WriteLine("");
    }
    public static bool TestCallWithBooleanFunction(bool B)
    {
        bool _TestCallWithBooleanFunction = false;
        _TestCallWithBooleanFunction = true;
        return _TestCallWithBooleanFunction;
    }
    // Also have Property in a comment
    public static List<string> testFunctionWithPropertyInName()
    {
        List<string> _testFunctionWithPropertyInName = null;
        _testFunctionWithPropertyInName = new List<string>();
        return _testFunctionWithPropertyInName;
    }
    public static void TestPrivateLocalFunctionCall()
    {
        PrivateLocalFunctionCall();
        PrivateLocalFunctionCall();
    }
    private static void PrivateLocalFunctionCall()
    {
        // empty
    }
    // This will only be readable if the file converts with correct braces.
    public static bool TestFileFinishesWell()
    {
        bool _TestFileFinishesWell = false;
        _TestFileFinishesWell = true;
        return _TestFileFinishesWell;
    }
    public static void VB6FileAccess()
    {
        int F = 0;
        dynamic ReadResult = null;
        string ReadResult2 = "";
        int ReadResult3 = 0;
        F = FreeFile();
        FileOpen(F, "C:\\abc.txt", VBFileMode("Binary")); // TODO: (VERIFY) Verify File Access: Open __S1 For Binary As #F
        FileGet(F, ReadResult); // TODO: (VERIFY) Verify File Access: Get #F, , ReadResult
        FilePut(F, ReadResult); // TODO: (VERIFY) Verify File Access: Put #F, , ReadResult
        Input(F, ReadResult); // TODO: (VERIFY) Verify File Access: Input #F, ReadResult
        ReadResult = LineInput(F); // TODO: (VERIFY) Verify File Access: Line Input #F, ReadResult
        FileClose(F); // TODO: (VERIFY) Verify File Access: Close #F
                      // '''''''''''''
    }
    public static frm getFrm()
    {
        frm _getFrm = null;
        _getFrm = frm.instance;
        return _getFrm;
    }
    public static void VerifyWith()
    {
        // Converted WITH statement: With frm // Should be permitted.
        frm.instance.Title = frm.instance.Title;
        frm.instance.Top = 0;
        // With getFrm() // TODO (not supported): Expression used in WITH.  Verify result: getFrm()
        dynamic __withVar675 = getFrm(); ; // Should NOT be permitted.
        __withVar675.Caption = "";
        __withVar675.Top = 0;
        // Converted WITH statement: With frm // Should be permitted.
        // With .txtFile // TODO (not supported): Expression used in WITH.  Verify result: .txtFile
        dynamic __withVar436 = frm.instance.txtFile; ; // Should NOT be permitted
        __withVar436.Text = "";
        __withVar436.Top = 0;
    }

}
