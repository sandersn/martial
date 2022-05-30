		Martial
	The Letter Frequency Analyzer
	     version 0.03
	http://spoony-bard.com/zackman/martial/martial.htm

These are the simple instructions*:
	To start: Click Open and choose to the file you want. Currently all extensions show by 
default. Of course, recommended are script files that are ready to 
insert. Click Open again to open another file. Click Reset if you change your mind.

	Then: In the Sample Length box, type the length of string you want to look for. 
You can enter any number of sample lengths or ranges using this format:
1    (normal)
1-20 (range)
-5   (default range from 1 to x)
1,5  (commas)
1,2-3,4,8,9,10-20  (everything)


	Next: Enter the number of results you want to see in the Number of Results box.
Enter a number or Martial will default to 25 to spite you. Do not enter negative numbers or you
will not see any results.

	Finally: Click Analyze and watch the cute little address counter spin away. 
If you get tired of the analyzation, you can always click Cancel and try another day. 
The file read can take a long time because the computation is fairly intensive. However, 
with the improved engine most lengths less than 5 should run fairly fast.
For an example of large lengths, 3-12 takes 9 minutes on a 45K file on my P3/600.

	Enjoy: The results will print in the following format:
Filename list
File length	Bytes actually counted(only matters with comments enabled)
Number of unique samples (for example, a typical 4K text file only has about 70 to 80 of length 1)
The sample lengths requested
The elapsed time
The samples, sorted most space first in this format
<Rank>: total space = <totalspace>; hits = <hits>; string = <string>; hex = <hex>; len = <length>

	Extras:
Output File: Martial will write its results to a file, which is helpful when you need LOTS of results.
A standard Win32 text box can only hold 32K of text, so anything over 600~700 results is truncated.
An output file gets around this limitation.

Output Table: Martial will write its results to a file in thingy table format (FF=text). Thingy only
supports double byte values, but newer editors support up to any length. You must, however,
specify a range for the results to be written to so that Martial doesn't override any of your matches.

Beep When Finished: Martial emits an ear piercing Standard Windows Beep when it is done
analysing your file.


Ignore Spaces: Instructs Martial to cut matches with spaces from the final results
Ignore Tabs: Instructs Martial to cut matches with tabs from the final results
Ignore Returns: Instructs Martial to cut matches with returns from the final results
Exclude Table: Instructs Martial to cut matches that are already in the thingy formatted table
that you specify.
WARNING:These matches still count towards the total number of matches so if you use these options,
you might want to crank up the number of matches a few above the number you really want.

Use Comments:
When comments are enabled, Martial skips all bytes within the comments. Alternatively, you can
check Read Only *Inside* Comments, and Martial will read only inside comments. Use this feature
to divide a script into two pieces. For example, one piece could be the Japanese and partial
English translation, while the other piece could be the finished English translation.
The comment styles offered are C, C++, Thingy, and custom. Obviously custom is the most useful.
Custom comment tags can be any length; leave the end comment tag blank for new line.


System requirements:
Win32 system
Visual Basic runtimes(free!)version 5 and 6. 5 is the program, 6 is the updated common dialog I use.
You are on your own to find both of these. Try Dell for the first, Microsoft for the second.
But you probably have both already anyway.

by ZackMan
reachable at sandersn@hotmail.com
http://spoony-bard.com/zackman

version history:
0.01 First release:
Releasing a very early version because of interest. Necrosaro has some interesting ideas
that could make this very much more useful.
0.02:
Added many new features: Comments, multiple parse lengths, improved match sorting, full
context sensitive rollover help.
0.03:
Revamped the GUI to be wizard style and added many new features:
Optimised search
Visual completion bar
Added filetypes and error checking to various dialogs
Ability to exclude values matching those already present in a thingy format table
Ability to write report to a file
Ability to write values to a file in thingy table format
Added beep on finish option
New report format
Multiple file batching

Martial is uses GPL v3.

*If you don't know what a letter frequency analyzer is useful for, you probably won't
be able to make sense of the rest of the document. Or Martial itself for that matter.
Besides, with the new wizard style interface, most people shouldn't have to read the readme
unless it's for fun anyway.
