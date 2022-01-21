using System;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modDirStack;
using static modShell;
using static modUtils;



static class modGit
{
    // Module to provide 'git' access form Immediate window.
    // Usage:  Type something like:  git __S1
    // Note:  `git` is a function.  Everything after that is passed as a single string.  To put quotes in that, double them.  e.g.
    // git __S1
    // Also note that you cannot do interactive git commands, such as `git commit -p`
    // For simple commands, such as status and push, the constants below allow the general syntax:  `git push`, `git status`.  Just dont be
    // confused and think this is a terminal.  It's still just a VB6 function running, and it only takes VB6 arguments.
    public const string Status = "status ";
    public const string St = "status ";
    public const string Commit = "commit -m ";
    public const string Push = "push ";
    public const string Pull = "pull ";
    public const string Branch = "branch ";
    public const string BR = "branch ";
    public const string Stash = "stash";
    public const string CheckOut = "checkout ";
    private static string GitFolder()
    {
        string _GitFolder = "";
        _GitFolder = AppContext.BaseDirectory + "\\";
        return _GitFolder;
    }
    public static string GitCmd(string C, bool NoOutput = false, bool HideCommand = false)
    {
        string _GitCmd = "";
        string ErrSt = "";
        PushDir(GitFolder());
        if (!HideCommand) GitOut("$ " + C);
        _GitCmd = RunCmdToOutput(C, out ErrSt);
        PopDir();
        if (!NoOutput) GitOut(_GitCmd);
        if (ErrSt != "") GitOut("ERR: " + ErrSt);
        return _GitCmd;
    }
    private static bool GitOut(string Msg)
    {
        bool _GitOut = false;
        Msg = Trim(Msg);
        while ((Left(Msg, 1) == vbCr || Left(Msg, 1) == vbLf)) { Msg = Mid(Msg, 2); }
        if (Len(Msg) > 0) Console.WriteLine(Msg);
        return _GitOut;
    }
    public static bool Git(string C)
    {
        bool _Git = false;
        if (LCase(Left(C, 4)) != "git ") C = "git " + C;
        GitCmd(C);
        _Git = true;
        return _Git;
    }
    public static void GitConf(string vName = "", string vEMail = "", bool Clear = false)
    {
        if (!IsIDE()) return;
        GitCmd("git config --unset user.name", true, true);
        GitCmd("git config --unset user.email", true, true);
        if (Clear)
        {
            GitCmd("git config --unset --global user.name", true);
            GitCmd("git config --unset --global user.email", true);
        }
        else if (vName == "" || vEMail == "")
        {
            Console.WriteLine("user.name=" + Trim(Replace(Replace(GitCmd("git config --global user.name", true, true), vbCr, ""), vbLf, "")));
            Console.WriteLine("user.email=" + Trim(Replace(Replace(GitCmd("git config --global user.email", true, true), vbCr, ""), vbLf, "")));
        }
        else
        {
            // GitCmd __S1, True
            // GitCmd __S1, True
            GitCmd("git config --global user.name " + vName);
            GitCmd("git config --global user.email " + vEMail);
        }
    }
    public static bool GitPull(bool withReset = true)
    {
        bool _GitPull = false;
        if (!IsIDE()) return _GitPull;
        // If withReset Then GitReset
        if (withReset)
        {
            GitCmd("git stash");
            GitCmd("git checkout master");
        }
        GitCmd("git pull -r");
        if (MsgBox("Restarting IDE in 5s...", vbOKCancel) == vbCancel) return _GitPull;
        // RestartIDE
        _GitPull = true;
        return _GitPull;
    }
    public static string GitStatus()
    {
        string _GitStatus = "";
        if (!IsIDE()) return _GitStatus;
        _GitStatus = GitCmd("git status");
        return _GitStatus;
    }
    public static string GitVersion()
    {
        string _GitVersion = "";
        if (!IsIDE()) return _GitVersion;
        _GitVersion = GitCmd("git --version");
        return _GitVersion;
    }
    public static bool HasGit()
    {
        bool _HasGit = false;
        if (!IsIDE()) return _HasGit;
        _HasGit = GitVersion() != "";
        return _HasGit;
    }
    public static bool GitReset(bool Hard = false, bool toMaster = false)
    {
        bool _GitReset = false;
        if (!IsIDE()) return _GitReset;
        if (!Hard)
        {
            GitCmd("git checkout -- .");
            if (toMaster) GitCmd("git checkout master -f");
        }
        else
        {
            if (toMaster) GitCmd("git checkout master -f");
            GitCmd("git reset --hard");
            GitCmd("git pull -r --force");
        }
        // RestartIDE
        _GitReset = true;
        return _GitReset;
    }
    public static bool GitPush(string Committer, string CommitMessage)
    {
        bool _GitPush = false;
        if (!IsIDE()) return _GitPush;
        GitCmd("git add .");
        GitCmd("git status");
        // If MsgBox(__S1, vbOKCancel + vbQuestion + vbDefaultButton1, __S2, , , 10) = vbCancel Then
        // GitCmd __S1
        // GitCmd __S1
        // Exit Function
        // End If
        GitCmd("git commit -m \"" + CommitMessage + "\"");
        GitCmd("git pull -r");
        // If MsgBox(__S1, vbOKCancel + vbQuestion + vbDefaultButton1, __S2, , , 10) = vbCancel Then
        // GitCmd __S1
        // GitCmd __S1
        // Exit Function
        // End If
        GitCmd("git push", true);
        GitCmd("git status");
        GitOut("GitPush Complete.");
        // If MsgBox(__S1, vbYesNo + vbExclamation + vbDefaultButton2, __S2, , , 10) = vbYes Then
        // GitProgress True
        // GitConf Clear:=True
        // GitProgress
        // End If
        _GitPush = true;
        return _GitPush;
    }
    public static void GitLog(int CharLimit = 3000)
    {
        string Res = "";
        if (!IsIDE()) return;
        Res = GitCmd("git log", true);
        Res = Left(Res, CharLimit);
        Console.WriteLine(Res);
    }
    public static bool GitCommits()
    {
        bool _GitCommits = false;
        GitCmd("git log --pretty=format:\"%h - %an, %ar : %s\" -10");
        _GitCommits = true;
        return _GitCommits;
    }
    public static bool GitRemoteBranches()
    {
        bool _GitRemoteBranches = false;
        GitCmd("git branch --remote --list");
        _GitRemoteBranches = true;
        return _GitRemoteBranches;
    }

}
