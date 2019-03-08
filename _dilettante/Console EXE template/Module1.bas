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

Private Type COORD
        x As Integer
        y As Integer
End Type

Private Type SMALL_RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
        dwSize As COORD
        dwCursorPosition As COORD
        wAttributes As Integer
        srWindow As SMALL_RECT
        dwMaximumWindowSize As COORD
End Type

Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long

Private Const FOREGROUND_BLUE = &H1     '  text color contains blue.
Private Const FOREGROUND_GREEN = &H2     '  text color contains green.
Private Const FOREGROUND_INTENSITY = &H8     '  text color is intensified.
Private Const FOREGROUND_RED = &H4     '  text color contains red.
'For SetConsoleMode (input)
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
'For SetConsoleMode (output)
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

Public Const STD_OUTPUT_HANDLE = -11&

Private Allocated As Boolean
Private hOutput As Long

Private Sub Main()
    Dim tmp As String
    
    Setup 'Omit for Console Subsystem.

    Dim scrbuf      As CONSOLE_SCREEN_BUFFER_INFO

    hOutput = GetStdHandle(STD_OUTPUT_HANDLE)
    GetConsoleScreenBufferInfo hOutput, scrbuf
    
    With New Scripting.FileSystemObject
        Set SIn = .GetStandardStream(StdIn)
        Set SOut = .GetStandardStream(StdOut)
    End With

    SOut.WriteLine "Enter text to echo:"
    tmp = SIn.ReadLine()
    SOut.WriteLine tmp
    
    SetConsoleTextAttribute hOutput, FOREGROUND_BLUE Or FOREGROUND_INTENSITY
    WriteToConsole "Blue Color !!" & vbCrLf
    
    'Change the text color to yellow
    SetConsoleTextAttribute hOutput, FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_INTENSITY
    WriteToConsole "Yellow Color !!" & vbCrLf
    
    'Restore the previous text attributes.
    SetConsoleTextAttribute hOutput, scrbuf.wAttributes
    
    TearDown 'Omit for Console Subsystem.
End Sub

'The following function writes the content of sText variable into the console window:
Private Function WriteToConsole(sText As String) As Boolean
    Dim lWritten            As Long
    
    If WriteFile(hOutput, ByVal sText, Len(sText), lWritten, ByVal 0) = 0 Then
        WriteToConsole = False
    Else
        WriteToConsole = True
    End If
End Function

Private Function ConsoleRead() As String
    Dim sUserInput As String * 256
    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
    'Trim off the NULL charactors and the CRLF.
    ConsoleRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function

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


