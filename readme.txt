
These are some common libraries I use in my code.
All C libraries are vb6 compatiable (stdcall)

pe_lib   - vb6 ActiveX DLL
           PE file format parsing code

proc_lib - vb6 ActiveX DLL
           Process Enumeration/Manipulation

globals  - vb6 Activex DLL
           Global Multiuse class with common lib functions

olly_dll - C DLL  
           ASM/DISASM engine 
           GPL Copyright (C) 2001 Oleh Yuschuk

hooklib  - C DLL
	   Detours style hooking engine 
           extended version of Daniel Pistelli's ntcore.com hook engine

zlib     - C DLL
           ZLIB compression library - win2k+ compatiable
           Jean-loup Gailly & Mark Adler

WinGraphviz - is a single DLL COM object directly usable from vb6.
	      It uses the code from Graphviz which is public and 
              freely available. Sample VB project included.

htmlViewer_lite - vb6 user control, lite weight htmlViewer based on rich 
                  text box control

vb6_utypes - adds support for unsigned math operations w/ ints and longs

kroolCtls - winapi replacements for mscomctl controls by kroll:
              listview, rtf, progressbar, tabstrip, treeview, + ipaddress
              see readme for more details (some mods to rtf planned)

gnu_whois - cmdline whois app, was going to convert to dll, but its fine
            as is, just capture output via pipe. 93k

filterList - simple listview with filter textbox for easy searching

anchor - very easy form element anchoring (automatic control resizing w/form)