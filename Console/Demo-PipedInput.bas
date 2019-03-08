Attribute VB_Name = "MDemoPipedInput"
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

Public Sub Main()
   Dim sData As String
   Dim sMessage As String
   
   ' Required in all MConsole.bas supported apps!
   Con.Initialize
   
   ' Check to see if we have any waiting input.
   If Con.Piped Then
      ' Slurp it all in a single stream.
      sData = Con.ReadStream()
      ' Just to prove we did it, place text on clipboard.
      Clipboard.Clear
      Clipboard.SetText sData
      ' Write some debugging information.
      Con.DebugOutput "Wrote " & CStr(Len(sData)) & " characters to clipboard."
   Else
      sMessage = "No redirection detected; nothing to read?"
      ' Send error condition to Standard Error!
      Con.WriteLine sMessage, True, conStandardError
      Con.DebugOutput sMessage
      ' Set an exit code appropriate to this error.
      ' Application MUST BE COMPILED TO NATIVE CODE to
      ' avoid a GPF in the runtime!!!
      Con.ExitCode = 1
   End If
End Sub

