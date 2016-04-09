Title: Fastest, safest subclasser, no module!
Description: *** 
Update: See my new submission here...
 http://www.exhedra.com/vb/scripts/ShowCode.asp?txtCodeId=42918&lngWId=1
If you do want the original zip then email me at Paul_Caton@hotmail.com
***
cSuperClass.cls is i believe the fastest, safest compiled in window subclasser around. Speed: The WndProc is executed entirely in run-time dynamically generated machine code. The class only calls back on messages that you choose. Safety: So far I've not been able to crash the IDE by pressing the end button or with the End statement. Flexible: The programmer can choose between filtered mode (fastest) and all messages mode. In filtered mode the user decides which windows messages they're interested in and can individually specify whether the message is to callback after default processing or before. Before mode additionally allows the programmer to specify whether or not default processing is to be performed subsequently.
No module: AFAIK this is the only subclasser ever to eschew the use of a module. So how do I get the address of the WndProc routine? Simple, the dynamically generated machine code lives in a byte array; you can get its address with the undocumented VarPtr function. The real magic in cSuperClass.cls is getting from the WndProc to the callback interface routine using ObjPtr against the owning Form/UserControl, see the assembler .asm model file included in the zip. Speaking of which... it may well be the case that my assembler is sub-optimal. Any experts out there willing to take a look? I thought I had a nifty/dirty stack trick working for a while but it didn't pan out. Should work with VB5 if VarPtr & ObjPtr were in that release? Sample project included. Regards.
This file came from Planet-Source-Code.com...the home millions of lines of source code
You can view comments on this code/and or vote on it at: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=37102&lngWId=1

The author may have retained certain copyrights to this code...please observe their request and the law by reviewing all copyright conditions at the above URL.
