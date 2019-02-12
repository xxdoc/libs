Attribute VB_Name = "Module1"
Option Explicit
'if the user hits the console window X in the IDE whole IDE is terminated
'note vbp file has had linker options manually added so it compiles as console no addin needed
'
'[VBCompiler]
'LinkSwitches=/SUBSYSTEM:CONSOLE

Public SIn As Scripting.TextStream 'Reference to Microsoft Scripting Runtime.
Public SOut As Scripting.TextStream

'--- Only required for testing in IDE or Windows Subsystem ===
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long

Private Allocated As Boolean


Private Sub Main()
    Dim tmp As String
    
    Setup 'Omit for Console Subsystem.

    With New Scripting.FileSystemObject
        Set SIn = .GetStandardStream(StdIn)
        Set SOut = .GetStandardStream(StdOut)
    End With

    SOut.WriteLine "Enter text to echo:"
    tmp = SIn.ReadLine()
    SOut.WriteLine tmp
    
    TearDown 'Omit for Console Subsystem.
End Sub

Private Sub Setup()
    Dim Title As String

    Title = Space$(260)
    If GetConsoleTitle(Title, 260) = 0 Then
        AllocConsole
        Allocated = True
    End If
End Sub

Private Sub TearDown()
    If Allocated Then
        SOut.Write "Press enter to continue..."
        SIn.ReadLine
        FreeConsole
    End If
End Sub
'--- End testing ---------------------------------------------


