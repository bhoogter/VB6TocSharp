# VB6 To C#

A VB6 based VB6 -> C# 2017 converter.

## Design Considerations

- Simple - Not designed to do a 100% conversion.  Just maybe an 80% - 90% of the grunt work.
- VB6 Based - Because, why not?  You have to have a working VB6 compiler if you're converting FROM vb6 anyway.
- Custom - This was created for a personal project, and hence, is specifically tailored for our use case.  But, there isn't any reason why someone couldn't invesitgate the logic and tweak it for any of their own issues.
- Opportunistic - This code heavily relies on relative uniformity of the VB6 IDE:
    - Spacing is relatively consistent because the IDE enforces it.
    - Keyword capitalization can be guaranteed.
- Non-assuming - It makes the assumption that the code compiled while in VB, so it doesn't assume that reference it can't resolve aren't going to be found.
- C# 2017 - This is a late-comer.  There has never been a freeware solution for VB6 -> C#, and now that VB.NET is more or less discontinued, why not?
