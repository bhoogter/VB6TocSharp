# VB6 To C#

A VB6 based VB6 -> C# 2017 converter.

## Usage

Free to use.  Free to fork.  Free to contribute.  Free to ask about.  Free to sell.  Free to sell under your own name...  Free to do just about anything except say I can't (See [LICENSE](https://github.com/bhoogter/VB6TocSharp/blob/master/LICENSE.txt)).

## Quick Start

Open the file `prj.vbp`, start the program.  Enter some config values, and convert a single file.

## Requirements

- Visual Studio supporting some relatively modern version of C#.  Or alternate.
- [Visual Basic Power Packs 2005 (v3.0)](https://www.microsoft.com/en-us/download/details.aspx?id=25169)
    - Allows use of Standard VB functions like `Mid`, `Trim`, `Abs`, `DateDiff`, etc, directly in C# code.
    - Ensures 99.9% compatibility with VB6 functionality (except for `Format`...), without the need for a 3rd party, black-box library (it's from MS, so it's a 1st party black box).
    - Easy to iterate off of once it's converted and up-and-running.

## Instructions

Please see the [wiki](https://github.com/bhoogter/VB6TocSharp/wiki) for more information on usage.

## Design Considerations

- Simple - Not designed to do a 100% conversion.  Just maybe an 80% - 90% of the grunt work.
- VB6 Based - Because, why not?  You have to have a working VB6 compiler if you're converting FROM vb6 anyway.
- Custom - This was created for a personal project, and hence, is specifically tailored for our use case.  But, there isn't any reason why someone couldn't invesitgate the logic and tweak it for any of their own issues.
- Opportunistic - This code heavily relies on relative uniformity of the VB6 IDE:
    - Spacing is relatively consistent because the IDE enforces it.
    - Keyword capitalization can be guaranteed.
    - We take advantage of the Microsoft Power Packs, and do NOT need to convert most of the core VB6 statements.  Further, you can continue to USE statements like `DateDiff`, `Left`, `Trim` as you would in VB.  Or, if you prefer, begin to migrate away from them AFTER conversion.  We simply pull in Microsoft's library for maximum compatibility, and hence, do not have a large string replace library, nor do we rely as heavily as some converters do on our own DLLs or libraries (we generate a few for ease of syntax, but the end result is pure C# code).
- Non-assuming - It makes the assumption that the code compiled while in VB, so it doesn't assume that reference it can't resolve aren't going to be found.
- C# 2017 - This is a late-comer.  There has never been a freeware solution for VB6 -> C#, and now that VB.NET is more or less discontinued, why not?

## Down-sides

- This will not produce code that will compile in its generated form.  The last mile is simply the most expensive, and it seemed more expedient to make something get most of the way, and finish the course manually.
- This isn't the most customizable solution.  Unless, of course, you want to dive into a little source code on the converter.  But, that's why its available.

## Pluses

- Do the whole thing or just one file at a time.
- It's free.
- You have the source.
- It's a lot better than doing it all by hand.
- It will give you a good insight into what's going on, without having to ALL of the manual effort to do a simple conversion.
- Not a fast conversion, but a strightforward one.  Inspect functions such as `ConvertSub` or `ConvertPrototype`.
- Allows inspection of how something is being converted.  Don't like the output?  Change it.

## Extras

- An albeit slow, but useful VB6 code linter.  Root out as much tech debt before even beginning the process.
- VB6 form to XAML

## Contact

- If you do have any questions, concerns, or simply would like some quick pointers, feel free to open an Issue.  I can't guarantee much, but I do try!
