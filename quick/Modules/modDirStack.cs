using Microsoft.VisualBasic;
using static Microsoft.VisualBasic.Conversion;
using static Microsoft.VisualBasic.FileSystem;



static class modDirStack
{
    public static Collection DirStack = new Collection(); // TODO: (NOT SUPPORTED) Dimmable 'New' not supported on variable declaration.  Instantiated only on declaration.  Please ensure usages
    public static string PushDir(string NewDir, bool doSet = true)
    {
        string _PushDir = "";
        // ::::PushDir
        // :::SUMMARY
        // :Basic Directory Stack - Push cur dir to stack and CD to parameter.
        // :::DESCRIPTION
        // :1. Push Current Dir to stack
        // :2. CD to new folder.
        // :::PAREMETERS
        // : - sNewDir - String - Directory to CD into.
        // : - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
        // :::RETURNS
        // :Returns current directory.
        // :::SEE ALSO
        // : PopDir, PeekDir
        int N = 0;
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (DirStack == null)
        {
            DirStack = new Collection();
            DirStack.Add(0, "n");
        }
        N = Val(DirStack.Item("n")) + 1;
        DirStack.Remove("n");
        DirStack.Add(N, "n");
        DirStack.Add(CurDir, "_" + N);
        if (doSet) ChDir(NewDir);
        _PushDir = CurDir;
        return _PushDir;
    }
    public static string PopDir(bool doSet = true)
    {
        string _PopDir = "";
        // ::::PopDir
        // :::SUMMARY
        // :Remove to dir from stack.  Error Safe.  Generally to change current directory.
        // :::DESCRIPTION
        // :1. Pop Dir from stack.
        // :2. CD to dir.
        // :::PAREMETERS
        // : - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
        // :::RETURNS
        // :Returns directory popped.
        // :::SEE ALSO
        // : PopDir, PeekDir
        int N = 0;
        string V = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (DirStack == null) return _PopDir;
        N = Val(DirStack.Item("n"));
        _PopDir = DirStack.Item("_" + N);
        if (N > 1)
        {
            N = N - 1;
            DirStack.Remove("n");
            DirStack.Add(N, "n");
        }
        else
        {
            DirStack = null;
        }
        if (doSet) ChDir(_PopDir);
        return _PopDir;
    }
    public static string PeekDir(bool doSet = true)
    {
        string _PeekDir = "";
        // ::::PeekDir
        // :::SUMMARY
        // :Return directory on top of stack without removing it.  Generally to change current directory.
        // :::DESCRIPTION
        // :1. Push Current Dir to stack
        // :2. CD to new folder.
        // :::PAREMETERS
        // : - [doSet] = True - Boolean - Pass FALSE if you don't want to change current directory.
        // :::RETURNS
        // :Returns top stack item (without removing it from stack).
        // :::SEE ALSO
        // : PopDir, PeekDir
        int N = 0;
        string V = "";
        // TODO: (NOT SUPPORTED): On Error Resume Next
        if (DirStack == null) return _PeekDir;
        N = Val(DirStack.Item("n"));
        _PeekDir = DirStack.Item("_" + N);
        if (doSet) ChDir(_PeekDir);
        return _PeekDir;
    }

}
