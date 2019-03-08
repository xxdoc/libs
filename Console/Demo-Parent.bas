Attribute VB_Name = "MDemoParent"
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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Main()
   ' No avoiding creation of console window, so
   ' we need to initialize our handler and deal.
   ' Determine whether we were launched within a console
   ' or via some other Explorer-like means.
   If Con.Initialize = conLaunchExplorer Then
      ' They double-clicked us in Explorer, so respond with
      ' Usage instructions via a MsgBox.
      Con.Visible = False
      MsgBox Usage(), vbOKOnly, "Startup from Explorer detected..."
      
   Else
      ' Normal console operation, even if running within the IDE!
      If Con.LaunchMode = conLaunchConsole Then
         Con.WriteLine "Startup from console detected..."
      ElseIf Con.LaunchMode = conLaunchVBIDE Then
         Con.WriteLine "Startup from within VB IDE detected..."
      End If
      Con.WriteLine "======================================"
      Con.WriteLine "Usage:"
      Con.WriteLine Usage()
   End If
End Sub

Private Function Usage() As String
   Usage = _
      "This is where you would supply your users with instructions" & vbCrLf & _
      "on how to use your console application. You could offer a" & vbCrLf & _
      "list of command line switches, for instance."
End Function

