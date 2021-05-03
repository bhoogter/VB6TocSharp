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


static class modGit {
// Option Explicit
public const string Status = "status ";
public const string ST = "status ";
public const string Commit = "commit -m ";
public const string Push = "push ";
public const string Pull = "pull ";
public const string Branch = "branch ";
public const string BR = "branch ";
public const string Stash = "stash";
public const string CheckOut = "checkout ";


private static string Gitfolder() {
  string Gitfolder = "";
  Gitfolder = App.Path + "\\";
  return Gitfolder;
}

private static string GitCmd(string C, bool NoOutput= false, bool HideCommand= false) {
  string GitCmd = "";
  string ErrSt = "";

  PushDir(Gitfolder);
  if (!HideCommand) {
    GitOut("$ " + C);
  }
  GitCmd = RunCmdToOutput(C, ref ErrSt);
  PopDir();
  if (!NoOutput) {
    GitOut(GitCmd);
  }
  if (ErrSt != "") {
    GitOut("ERR: " + ErrSt);
  }
  return GitCmd;
}

private static bool GitOut(string Msg) {
  bool GitOut = false;
  Msg = Trim(Msg);
  while((Left(Msg, 1) == vbCr || Left(Msg, 1) == vbLf)) {
    Msg = Mid(Msg, 2);
  }
  if (Len(Msg) > 0) {
    Console.WriteLine(Msg);
  }
  return GitOut;
}

public static bool Git(string C) {
  bool Git = false;
  if (LCase(Left(C, 4)) != "git ") {
    C = "git " + C;
  }
  GitCmd(C);
  Git = true;
  return Git;
}

public static dynamic GitConf(string vName= "", string vEMail= "", bool Clear= false) {
  dynamic GitConf = null;
  if (!IsIDE()) {
    return GitConf;

  }

  GitCmd("git config --unset user.name", true, true);
  GitCmd("git config --unset user.email", true, true);
  if (Clear) {
    GitCmd("git config --unset --global user.name", true);
    GitCmd("git config --unset --global user.email", true);
  } else if (vName == "" || vEMail == "") {
    Console.WriteLine("user.name=" + Trim(Replace(Replace(GitCmd("git config --global user.name", true, true), vbCr, ""), vbLf, "")));
    Console.WriteLine("user.email=" + Trim(Replace(Replace(GitCmd("git config --global user.email", true, true), vbCr, ""), vbLf, "")));
  } else {
//    GitCmd "git config --unset --global user.name", True
//    GitCmd "git config --unset --global user.email", True
    GitCmd("git config --global user.name " + vName);
    GitCmd("git config --global user.email " + vEMail);
  }
  return GitConf;
}

public static bool GitPull(bool withReset= true) {
  bool GitPull = false;
  if (!IsIDE()) {
    return GitPull;

  }
//  If withReset Then GitReset
  if (withReset) {
    GitCmd("git stash");
    GitCmd("git checkout master");
  }

  GitCmd("git pull -r");
  if (MsgBox("Restarting IDE in 5s...", vbOKCancel) == vbCancel) {
    return GitPull;

  }
//  RestartIDE
  GitPull = true;
  return GitPull;
}

public static dynamic GitStatus() {
  dynamic GitStatus = null;
  if (!IsIDE()) {
    return GitStatus;

  }
  GitCmd("git status");
  GitStatus = true;
  return GitStatus;
}

public static bool GitReset(bool Hard= false, bool toMaster= false) {
  bool GitReset = false;
  if (!IsIDE()) {
    return GitReset;

  }
  if (!Hard) {
    GitCmd("git checkout -- .");
    if (toMaster) {
      GitCmd("git checkout master -f");
    }
  } else {
    if (toMaster) {
      GitCmd("git checkout master -f");
    }
    GitCmd("git reset --hard");
    GitCmd("git pull -r --force");
  }
//  RestartIDE
  GitReset = true;
  return GitReset;
}

public static bool GitPush(string Committer_UNUSED, string CommitMessage) {
  bool GitPush = false;
  if (!IsIDE()) {
    return GitPush;

  }

  GitCmd("git add .");
  GitCmd("git status");

//  If MsgBox("Continue with Commit?", vbOKCancel + vbQuestion + vbDefaultButton1, "git push", , , 10) = vbCancel Then
//    GitCmd "git stash clear"
//    GitCmd "Stash cleared."
//    Exit Function
//  End If

  GitCmd("git commit -m \"" + CommitMessage + "\"");
  GitCmd("git pull -r");

//  If MsgBox("Continue with Push?", vbOKCancel + vbQuestion + vbDefaultButton1, "git push", , , 10) = vbCancel Then
//    GitCmd "git stash clear"
//    GitCmd "Stash cleared."
//    Exit Function
//  End If

  GitCmd("git push", true);
  GitCmd("git status");

  GitOut("GitPush Complete.");

//  If MsgBox("Clear Credentials?", vbYesNo + vbExclamation + vbDefaultButton2, "Done?", , , 10) = vbYes Then
//    GitProgress True
//    GitConf Clear:=True
//    GitProgress
//  End If

  GitPush = true;
  return GitPush;
}

public static dynamic GitLog(int CharLimit= 3000) {
  dynamic GitLog = null;
  string Res = "";

  if (!IsIDE()) {
    return GitLog;

  }
  Res = GitCmd("git log", true);
  Res = Left(Res, CharLimit);
  Console.WriteLine(Res);
  return GitLog;
}

public static bool GitCommits() {
  bool GitCommits = false;
  GitCmd("git log --pretty=format:\"%h - %an, %ar : %s\" -10");

  GitCommits = true;
  return GitCommits;
}

public static bool GitRemoteBranches() {
  bool GitRemoteBranches = false;
  GitCmd("git branch --remote --list");
  GitRemoteBranches = true;
  return GitRemoteBranches;
}
}
