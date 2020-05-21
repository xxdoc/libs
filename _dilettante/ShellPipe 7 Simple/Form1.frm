VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Shell Program"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin Project1.ShellPipe SP 
      Left            =   2280
      Top             =   300
      _ExtentX        =   635
      _ExtentY        =   635
      ErrAsOut        =   0   'False
   End
   Begin VB.Label lblComplete 
      Caption         =   "Complete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Simple demonstration of the ShellPipe control.
'
'Here we attempt to run an external script via CScript, then
'we read lines of text from inputdata.txt and feed them to
'the script a line at a time.
'
'The script then replays each of those lines back to us
'via StdErr and we write those to errordata.txt, and it
'also echos back these lines reversed via StdOut and we
'write those to outputdata.txt, and then it closes the
'output streams and/or CScript.exe finishes.
'
'Then we clean up and write the return code from CScript to
'outputdata.txt, and we're done.
'

Private Sub cmdGo_Click()
    Dim SPResult As SP_RESULTS
    Dim TextLine As String
    
    cmdGo.Enabled = False

    On Error Resume Next
    Kill "outputdata.txt"
    Kill "errordata.txt"
    On Error GoTo 0
    Open "inputdata.txt" For Input As #1
    Open "outputdata.txt" For Output As #2
    Open "errordata.txt" For Output As #3
    
    SPResult = SP.Run("cscript test.vbs //nologo")
    Select Case SPResult
        Case SP_SUCCESS
            Do Until EOF(1)
                Line Input #1, TextLine
                SP.SendLine TextLine
            Loop
            SP.ClosePipe

        Case SP_CREATEPIPEFAILED
            MsgBox "Run failed, could not create pipe", _
                   vbOKOnly Or vbExclamation, _
                   Caption

        Case SP_CREATEPROCFAILED
            MsgBox "Run failed, could not create process", _
                   vbOKOnly Or vbExclamation, _
                   Caption
    End Select
End Sub

Private Sub SP_ChildFinished()
    Dim lngReturnCode As Long
    
    'Pick up any leftover output prior to child termination.
    If SP.ErrLength > 0 Then Print #3, SP.ErrGetData()
    If SP.Length > 0 Then Print #2, SP.GetData()
    
    lngReturnCode = SP.FinishChild(0)
    Close #3
    Print #2, "Program complete. Return code:"; lngReturnCode
    Close #2
    Close #1
    lblComplete.Visible = True
End Sub

Private Sub SP_DataArrival(ByVal CharsTotal As Long)
    With SP
        Do While .HasLine
            Print #2, .GetLine()
        Loop
    End With
End Sub

Private Sub SP_EOF(ByVal EOFType As SPEOF_TYPES)
    'Pick up any leftover output prior to EOF.
    If SP.Length > 0 Then Print #2, SP.GetData()
    
    Print #2, "*EOF on StdOut*"
End Sub

Private Sub SP_ErrDataArrival(ByVal CharsTotal As Long)
    With SP
        Do While .ErrHasLine
            Print #3, .ErrGetLine()
        Loop
    End With
End Sub

Private Sub SP_ErrEOF(ByVal EOFType As SPEOF_TYPES)
    'Pick up any leftover output prior to EOF.
    If SP.ErrLength > 0 Then Print #3, SP.ErrGetData()
    
    Print #3, "*EOF on StdErr*"
End Sub

Private Sub SP_Error(ByVal Number As Long, ByVal Source As String, CancelDisplay As Boolean)
    MsgBox "Error " & CStr(Number) & " in " & Source, _
           vbOKOnly Or vbExclamation, _
           Caption
    CancelDisplay = True
    SP.FinishChild 0
End Sub

