# VB6 To C#

A VB6 based VB6 -> C# 2017 converter.

## Usage

Free to use.  Free to fork.  Free to contribute.  Free to ask about.  Free to sell.  Free to sell under your own name...  Free to do just about anything except say I can't (See [LICENSE](https://github.com/bhoogter/VB6TocSharp/blob/master/LICENSE.txt)).

## Quick Start

1. Open the file `prj.vbp`, start the program (Requires the VB6 IDE).
1. Enter some config values into the `Config` form via the button.
1. Now that you have selected your project, click the `SCAN` button.  This helps the converter know the difference between Methods without parenthesis and variables/constants.  It also builds a full list of imports (which can be cleaned up in the VS 2019 IDE via ^K^E).
1. If you want, click `SUPPORT` to generate the basic project support structure.
    - Alternatively, the files `VBExtension.cs` and `VBConstants.cs` could be copied directly out of the project root and included somewhere else.
1. Then, enter a filename and click `Single File` to try to convert the file you enter beside it.

If you want to convert the entire project, simply click `ALL`, and it will do the scan, support generation, and entire project conversion off the bat.  

NOTE:  It might not be the fastest, it still requires a manual effort, but it's faster than doing it ALL manually!

## Updates 2021-12-01

The original version of the converter approached the problem in a block-by-block approach, separating every logical program unit into its own string and converting it on its own.  As a result, the converter made multiple passes and basically ran extremely slow.  After updating the linter to not use this approach, but simply to run through the code from top to bottom, it appeared evident that the converter could do the same thing with just as much accuracy.

So, released today is the v2 of the converter, along side v1, via radio buttons on the main form.  Feel free to mix and match, convert the whole project with one, and then individual files with the other.  Hopefully, with the two completely different approaches, one of the two will work out better than the other and will result in less work.  Again, they are still, both, only an 80-90% conversion.  There are many things you will have to do by hand and double-check (like loop bounds), but again, it sure beats changing all of the `&`s in your VB6 code to `+`s for C#.

## How Do I ...?

There are a lot of questions when it comes to conversion.  If you just want to know the way this converter deals with specific patterns, please see the [How Do I ...?](https://github.com/bhoogter/VB6TocSharp/wiki/How-Do-I-...%3F) page in our wiki.

Whether you use this converter or not, we give our solution to commonly encountered conversion puzzles.  Our solutions are quick, to the point, and don't generally use a lot of programming or context overhead.  While they may rely on our extension module, all of it is native C# code, and, generally, fairly similar to what you did int VB6.

## Requirements

### Converter requirements
- The converter runs in the VB6 IDE.  You know, the IDE of the program you're trying to convert.

### Converted program requirements
- Visual Studio supporting some relatively modern version of C#.  Or alternate.
- [Visual Basic Power Packs](http://go.microsoft.com/fwlink/?LinkID=145727&clcid=0x804)
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
- Universal imports - Imports EVERY code module, just like VB6 did automatically.  Make Visual Studio do the work of deciding which ones are used simply by optimizing imports after conversion.
- C# 2017 - This is a late-comer.  There has never been a freeware solution for VB6 -> C#, and now that VB.NET is more or less discontinued, why not?

## Known Issues (v1 only)

- Currently, the converter will often balk at a file that contains the word 'Property' anywhere in it (other than a property declaration).  While this is a pain and likely will be fixed, it was encountered towards the end of the projects usefulness (and hence not urgent on the repair list), and where it hindered progress, the variable that contained the word '...Property' was simply renamed temporarily to something like '...Prppty', and then changed back in the converted file.
    - This has been addressed in Version 2, however it may still appear in some cases in v1.

## Down-sides

- This will not produce code that will compile in its generated form.  The last mile is simply the most expensive to automate, and often best to do manually.  It seemed more expedient to make something get most of the way, and finish any edge cases or final conversion by hand.
- Limited UI customziation (but limitless code-based customization).  This isn't the most customizable solution.  Unless, of course, you want to dive into a little source code on the converter.  But, that's why its available.
- Unlinted output.  The resultant code is a mess, stylistically.  That's what a modern IDE is for.  All the bad formatting can be cleaned up with ^K^D.  Unused imports with ^K^E.  And, there's a lot of extra {'s and }'s you will probably want to delete.
- The converter currently is REALLY bad at loop bounds.  Sorry, it's one of the pitfalls of a VB6->C# conversion, and there isn't much logic into how it converts.  It's tedius, but do a project-wide search for all for loops and manually inspect the bounds.
- There is an extra method for every event.  One for the correct signature, one for the original signature.  In most cases, the redundancy is unnecessary, but it provided the easiest conversion.  These can be reduced to a single method in most cases (but not all, which is why I don't).

NOTE:  Most of the places the converter knows there will be a problem it will annotate the code with a `// TODO:` comment.  Be sure to thoroughly address each of these.

## Pluses

- It's free.
- You have the source (customize it, whatever).
- Do the whole thing or just one file at a time.
- It's a lot better than doing it all by hand.
- It will give you a good insight into what's going on, without having to ALL of the manual effort to do a simple conversion.
- Not the fastest conversion, but a strightforward one (but v2 just got a whole lot better).  Inspect functions such as `ConvertSub` or `ConvertPrototype`.  
    - But think about it...  You're looking to convert in a one-off, not run over and over during converted exectution.
- Allows inspection of how something is being converted.  Don't like the output?  Change it.
    - You can put a VB6 breakpoint anywhere you like and stop.  Additionally, just add a line like `If LineN = 387 Then Stop`, and the converter will stop right there.

## Extras

- A VB6 code linter. `?Lint`.  Root out as much tech debt before even beginning the process.
- VB6 form to XAML

## Contact

- If you do have any questions, concerns, or simply would like some quick pointers, feel free to open an Issue.  I can't guarantee much, but I do try!
