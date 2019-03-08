Attribute VB_Name = "MDemoExitCode"
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
   Dim nRet As Long
   
   ' Required in all MConsole.bas supported apps!
   Con.Initialize
   
   ' Tell user what's supposed to happen, if no command line.
   If Len(Command$) Then
      ' Simply attempt to coerce command line to a numeric value,
      ' and return that as the exitcode for this application.
      On Error Resume Next
         nRet = Val(Command$)
      On Error GoTo 0
   Else
      Con.WriteLine "This demo returns the numeric value of the command line as its ExitCode."
      Con.WriteLine "Example usage (employing retval.bat sample):"
      Con.WriteLine
      Con.WriteLine "   retval condemo3 42"
      Con.WriteLine
      ' Allow user to read usage info, if not started within console.
      If Con.LaunchMode = conLaunchExplorer Then
         Con.PressAnyKey
      End If
   End If
   
   ' Important Note: The VB runtime has been known to
   ' GPF if EXEs that set their exitcode aren't compiled
   ' to "native" mode.  For your own safety, this setting
   ' is strongly recommended!
   Con.ExitCode = nRet
End Sub

