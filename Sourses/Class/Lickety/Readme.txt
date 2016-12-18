http://www.vbforums.com/showthread.php?781599-VB6-Lickety-an-alternative-to-Slurp-n-Split

[VB6] Lickety - an alternative to Slurp'n'Split
We're in a time of PCs with vast amounts of RAM and little going on in the background, despite money being spent on multi-core CPUs. What was once considered a crude and naive practice of slurping entire text files into a String and then calling Split() to make an array is now actually touted by many programmers.

But much as the practice of Slurp'n'Split was impractical when far smaller amounts of RAM were available, it can break down when you deal with large files.

First you have the problem of an ANSI text file doubling in size in RAM as soon as your slurp it. Then you double that again by calling the Split() function, at least until you discard the original intact String. Both of these together conspire against you for the maximum file size you can process no matter how much RAM you bought on Mommy's credit card: VB6 programs can only use just so much due to 32-bit and other design limitations.

To add insult to injury, Split() was taken over almost without modification from VBScript and never designed for high performance on large amounts of data. That just wasn't one of its design goals.


So What's A Slurp'n'Splitter To Do?

VB5 didn't get a Split() function, and this led the fans at VBspeed to create several alternative equivalents. And of course being VBspeed they worked out some clever optimizations.


Lickety Class

My Lickety.cls was based on some of those Split() alternatives, adding in the reading of the file in chunks for better memory efficiency.

The performance is pretty good, and as soon as your input file is much over 10KB of ANSI text it quickly overtakes a Split() call in both performance and memory requirements. This advantage just grows with the size of the input file.

It has options to read ANSI or Unicode (UTF-16LE) files, optionally skip the Unicode BOM, and whatever line delimiter you choose.

It also takes care of discarding the "dangler" empty line at the end of the file, unlike a Split() call.

I haven't looked it over to be sure but it may work in VB5 as well if you work around the returning of the array (VB5 can't have array-valued functions).