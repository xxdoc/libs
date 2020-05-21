Attribute VB_Name = "Module1"
'small command line utility to calculate inflation using US CPI
'this can run as either a console app or standard exe
'if you want it to run as a console app you can use linktool to process
'inf.vbc file contains command to change subsystem to console on compilation
'https://github.com/dzzie/addins/tree/master/LinkTool

'Console functions: https://gist.github.com/xaprb/8492636

'Inflation calculator dollars startYear [endYear=current] -d -c
'     -d  diff mode
'     -c  add commas
'
'Example: inf.exe 160k 2006 -c

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_INPUT_HANDLE = -10&

Sub Main()
    Dim cpi As New CInflation
    Dim d, y1, y2
    Dim dd As Double
    Dim c
    Dim diffMode As Boolean
    Dim addCommas As Boolean
    Dim outval
    
    c = Command
    c = Trim(Replace(c, "  ", " "))
    
    If InStr(c, "-d") > 0 Then
        c = Trim(Replace(c, "-d", Empty))
        diffMode = True
    End If
    
    If InStr(c, "-c") > 0 Then
        c = Trim(Replace(c, "-c", Empty))
        addCommas = True
    End If
    
    'MsgBox Command
    args = Split(c, " ")
    
    If UBound(args) = 1 Then
        d = args(0)
        y1 = args(1)
        y2 = Year(Now)
    ElseIf UBound(args) = 2 Then
        d = args(0)
        y1 = args(1)
        y2 = args(2)
    Else
        WriteStdOut Replace("\nInflation calculator dollars startYear [endYear=current] -d -c\n     -d  diff mode\n     -c  add commas\n   Example: inf.exe 160k 2006 -c\n\n", "\n", vbCrLf)
        End
    End If
    
    If InStr(1, d, "k", vbTextCompare) > 0 Then d = Replace(d, "k", "000")
    
    'MsgBox Join(args, vbCrLf)
    'MsgBox Join(Array(d, y1, y2), vbCrLf)
    
    If Not cpi.calculate(d, y1, y2, dd) Then
        WriteStdOut cpi.errMsg
    Else
        If diffMode Then
            outval = dd - cpi.dollarsSanitized
        Else
            outval = dd
        End If
        mask$ = "###,###,###,###,###0"
        If addCommas Then outval = Format(outval, mask$)
        WriteStdOut outval
    End If
   
End Sub

Function ReadStdIn(Optional ByVal NumBytes As Long = -1) As String
    Dim StdIn As Long
    Dim Result As Long
    Dim Buffer As String
    Dim BytesRead As Long

    StdIn = GetStdHandle(STD_INPUT_HANDLE)
    
    If StdIn = 0 Then
        ReadStdIn = InputBox("Enter console input")
        Exit Function
    End If
    
    Buffer = Space$(1024)
    
    Do
        Result = ReadFile(StdIn, ByVal Buffer, Len(Buffer), BytesRead, ByVal 0&)
        If Result = 0 Then
            Err.Raise 1001, , "Unable to read from standard input"
        End If
        ReadStdIn = ReadStdIn & Left$(Buffer, BytesRead)
    Loop Until BytesRead < Len(Buffer)
    
End Function

Sub WriteStdOut(ByVal Text As String)
    Dim StdOut As Long
    Dim Result As Long
    Dim BytesWritten As Long

    StdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    
    If StdOut = 0 Then
        MsgBox Text 'not compiled as a console app
        Exit Sub
    End If
    
    Result = WriteFile(StdOut, ByVal Text, Len(Text), BytesWritten, ByVal 0&)
    
    If Result = 0 Then
        Err.Raise 1001, , "Unable to write to standard output"
    ElseIf BytesWritten < Len(Text) Then
        Err.Raise 1002, , "Incomplete write operation"
    End If
    
End Sub


