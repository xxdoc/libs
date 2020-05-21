Attribute VB_Name = "MDemoInteractive"
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
   Dim fColor As Long, bColor As Long
   Dim sCaption As String
   Dim i As Long, j As Long, n As Long
   Const Twirl As String = "\|/-"
   
   ' Required in all MConsole.bas supported apps!
   Con.Initialize
   
   ' Stash value(s) we'll later reset.
   bColor = Con.BackColor
   fColor = Con.ForeColor
   sCaption = Con.Title
   
   ' Read and write a simple response:
   If Con.Height < 50 Then Con.Height = 50
   Con.ForeColor = conGreenHi
   Con.WriteLine "What's your name? ", False
   Con.ForeColor = fColor
   sName = Con.ReadLine()
   Con.ForeColor = conGreenHi
   Con.WriteLine "Hello " & sName, False
   Con.Title = "Console Demo for " & sName
   
   ' Show what sort of colors are available.
   Con.WriteLine ".  Here's a map of potential color selections..."
   For i = 0 To 15
      Con.BackColor = bColor
      Con.ForeColor = fColor
      Con.WriteLine Right$(Space$(6) & CStr(i) & ": ", 6), False
      Con.BackColor = i
      For j = 0 To 15
         Con.ForeColor = j
         Con.WriteLine " " & Format$(j, "00") & " ", False
      Next j
      Con.WriteLine
   Next i
   Con.BackColor = bColor
   Con.WriteLine vbCrLf
   
   ' Display some informational text...
   Con.ForeColor = conRedHi
   Con.WriteLine "Press Ctrl-C or Ctrl-Break to see how", , , conAlignCentered
   Con.WriteLine "this console module can handle these events" & vbCrLf, , , conAlignCentered
   Con.WriteLine "Now simulating a long-running process..." & vbCrLf, , , conAlignCentered
   
   'Simulate a long-running process:
   Con.CurrentX = 0
   Con.CursorVisible = False
   Con.ForeColor = conYellowHi
   Con.WriteLine "Working.", False
   For i = 1 To 50000
      ' Do a small part of a large task.
      If i Mod 30 = 0 Then
         n = n + 1
         Con.WriteLine Mid$(Twirl, (n Mod 4) + 1, 1), False
         Con.CurrentX = Con.CurrentX - 1
         If i Mod 600 = 0 Then
            Con.WriteLine ".", False
         End If
      End If
      
      'See if a shutdown event was caught:
      If Con.Break Then
         Con.WriteLine vbCrLf
         Con.CursorVisible = True
         
         Select Case Con.ControlEvent
            '  Note: A Win95 bug prevents some events from signaling!
            '  http://support.microsoft.com/default.aspx?scid=kb;en-us;130717
            Case conEventControlC
               Con.WriteLine "Ctrl-C pressed."
            
            Case conEventControlBreak
               ' In the IDE, a Ctrl-Break will result in the
               ' VB debugger breaking in the MConsole module
               ' after the next time VB looks for input.
               Con.WriteLine "Ctrl-Break pressed."
            
            Case conEventClose
               ' Unsupported in Win95!
               Con.WriteLine "User pressed the Big X to close window."
            
            Case conEventLogoff, conEventShutdown
               ' Unsupported in Win95!
               Con.WriteLine "User shutting down or logging off!"
               
         End Select
          
         ' Termination event was caught, so we should bail out...
         Exit For
      End If
      
      ' Take a deep breath...
      Sleep
   Next i

   ' For this demo only, allow user to press a key
   ' before potentially seeing the window closed.
   ' Normally, you would want to get the heck out
   ' of Dodge on a Break event.
   Con.BackColor = bColor
   Con.ForeColor = fColor
   Con.Title = sCaption
   If Con.Compiled Then Con.PressAnyKey
End Sub

