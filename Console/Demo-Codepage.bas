Attribute VB_Name = "MDemoCodepage"
' *************************************************************************
'  Copyright ©2004 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
'  Redistributed - with full permission - on http://www.vbadvance.com
' *************************************************************************
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (Optional ByVal dwMilliseconds As Long = 1)

Public Sub Main()
   Dim sName As String
   Dim sCaption As String
   Dim i As Long, j As Long, n As Long
   Const Twirl As String = "\|/-"
   
   ' Required in all MConsole.bas supported apps!
   Con.Initialize
   
   ' Read and write a simple response:
   'Con.WriteLine "What's your name? ", False
   'sName = Con.ReadLine()
   
   Con.Title = "CodePage mapping for: Hélène"
   
   ' Show what sort of colors are available.
   Con.WriteLine "Here's a map of CodePage " & Con.CodePageO & " characters..."
   For i = 0 To 15
      For j = 0 To 15
         Con.WriteLine " " & Chr$(i * 16 + j) & " ", False
      Next j
      Con.WriteLine
   Next i
   
   ' For this demo only, allow user to press a key
   ' before potentially seeing the window closed.
   ' Normally, you would want to get the heck out
   ' of Dodge on a Break event.
   Con.Title = sCaption
   If Con.Compiled Then Con.PressAnyKey
End Sub

