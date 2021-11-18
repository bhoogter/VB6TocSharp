using System;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modDirStack;
using static modShell;
using static modUtils;


static class modGit
{
    // Option Explicit
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
        string GitFolder = "";
        GitFolder = AppDomain.CurrentDomain.BaseDirectory + "\\";
        return GitFolder;
    }

    public static string GitCmd(string C, bool NoOutput = false, bool HideCommand = false)
    {
        string GitCmd = "";
        string ErrSt = "";

        PushDir(GitFolder());
        if (!HideCommand)
        {
            GitOut("$ " + C);
        }
        GitCmd = RunCmdToOutput(C, ref ErrSt);
        PopDir();
        if (!NoOutput)
        {
            GitOut(GitCmd);
        }
        if (ErrSt != "")
        {
            GitOut("ERR: " + ErrSt);
        }
        return GitCmd;
    }

    private static bool GitOut(string Msg)
    {
        bool GitOut = false;
        Msg = Trim(Msg);
        while ((Left(Msg, 1) == vbCr || Left(Msg, 1) == vbLf))
        {
            Msg = Mid(Msg, 2);
        }
        if (Len(Msg) > 0)
        {
            Console.WriteLine(Msg);
        }
        return GitOut;
    }

    public static bool Git(string C)
    {
        bool Git = false;
        if (LCase(Left(C, 4)) != "git ")
        {
            C = "git " + C;
        }
        GitCmd(C);
        Git = true;
        return Git;
    }

    public static void GitConf(string vName = "", string vEMail = "", bool Clear = false)
    {
        if (!IsIDE())
        {
            return;

        }

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
            //    GitCmd "git config --unset --global user.name", True
            //    GitCmd "git config --unset --global user.email", True
            GitCmd("git config --global user.name " + vName);
            GitCmd("git config --global user.email " + vEMail);
        }
    }

    public static bool GitPull(bool withReset = true)
    {
        bool GitPull = false;
        if (!IsIDE())
        {
            return GitPull;

        }
        //  If withReset Then GitReset
        if (withReset)
        {
            GitCmd("git stash");
            GitCmd("git checkout master");
        }

        GitCmd("git pull -r");
        if (MsgBox("Restarting IDE in 5s...", vbOKCancel) == vbCancel)
        {
            return GitPull;

        }
        //  RestartIDE
        GitPull = true;
        return GitPull;
    }

    public static string GitStatus()
    {
        string GitStatus = "";
        if (!IsIDE())
        {
            return GitStatus;

        }
        GitStatus = GitCmd("git status");
        return GitStatus;
    }

    public static string GitVersion()
    {
        string GitVersion = "";
        if (!IsIDE())
        {
            return GitVersion;

        }
        GitVersion = GitCmd("git --version");
        return GitVersion;
    }

    public static bool HasGit()
    {
        bool HasGit = false;
        if (!IsIDE())
        {
            return HasGit;

        }
        HasGit = GitVersion() != "";
        return HasGit;
    }

    public static bool GitReset(bool Hard = false, bool toMaster = false)
    {
        bool GitReset = false;
        if (!IsIDE())
        {
            return GitReset;

        }
        if (!Hard)
        {
            GitCmd("git checkout -- .");
            if (toMaster)
            {
                GitCmd("git checkout master -f");
            }
        }
        else
        {
            if (toMaster)
            {
                GitCmd("git checkout master -f");
            }
            GitCmd("git reset --hard");
            GitCmd("git pull -r --force");
        }
        //  RestartIDE
        GitReset = true;
        return GitReset;
    }

    public static bool GitPush(string Committer_UNUSED, string CommitMessage)
    {
        bool GitPush = false;
        if (!IsIDE())
        {
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

    public static void GitLog(int CharLimit = 3000)
    {
        string Res = "";

        if (!IsIDE())
        {
            return;

        }
        Res = GitCmd("git log", true);
        Res = Left(Res, CharLimit);
        Console.WriteLine(Res);
    }

    public static bool GitCommits()
    {
        bool GitCommits = false;
        GitCmd("git log --pretty=format:\"%h - %an, %ar : %s\" -10");

        GitCommits = true;
        return GitCommits;
    }

    public static bool GitRemoteBranches()
    {
        bool GitRemoteBranches = false;
        GitCmd("git branch --remote --list");
        GitRemoteBranches = true;
        return GitRemoteBranches;
    }
}
