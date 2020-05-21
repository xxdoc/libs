=========================================
ShellPipe control version 7 demonstration
=========================================

ShellPipe is a VB6 internal UserControl that can be
used to run external programs and communicate with
them via StdIn, StdOut, and StdErr.

It is patterned somewhat on the Microsoft Winsock
control in terms of methods and events.

This version uses a Timer control internally to
ShellPipe to accomodate the limitations of Win9x
platforms.  An NT-only version might be created using
async I/O if desired.

A different type of polling timer should be used if
this UserControl is changed to a VB6 Class module.

It would also be possible to modify ShellPipe to be
line-oriented on output from the spawned external
process.  In other words an event could be triggered
when a full line of output was received from the
spawned process.

The demonstration project runs a WSH script that
accepts String values, then returns them reversed.


Demo Project Files
==================

Project1.vbp    Demo's project file.
Project1.vbw    Demo's project workspace file.
Form1.frm       Main program, form module.
ShellPipe.ctl   ShellPipe UserControl module. (1)
ShellPipe.ctx   ShellPipe Binary Resource file. (1)
SPBuffer.cls    SPBuffer Class module. (1)


Other Files
===========

test.vbs        WSH script run by the demo program.
inputdata.txt   Data fed to test.vbs by demo.
outputdata.txt  StdOut results from test.vbs,
                captured by demo program.
errordata.txt   StdErr results from test.vbs,
                captured by demo program.
ReadMe.txt      This brief writeup.

Note: Files outputdata.txt and errordata.txt are
      not present in the archive package, but
      created when the demo is run.


Using ShellPipe
===============

Copy the three files annotated with (1) above into
your program's project folder.  Then add the .cls
and .ctl files to your project from the VB6 IDE's
Project Explorer window. (Add|File...).


Caveats
=======

This will not work with a program like the FTP.exe
included with Windows 2000, XP, etc.  Programs of that
type do direct console device I/O operations instead
of using standard I/O streams to talk to the console.

