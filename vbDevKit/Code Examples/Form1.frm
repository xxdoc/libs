VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB Developers Kit Sample Code"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help Documentation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   5835
   End
   Begin VB.CommandButton cmdClsCmdLine 
      Caption         =   "clsCmdLine Example"
      Height          =   555
      Left            =   1860
      TabIndex        =   6
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton cmdClsCmnDlg 
      Caption         =   "clsCmnDlg Example"
      Height          =   495
      Left            =   3660
      TabIndex        =   5
      Top             =   1620
      Width           =   2955
   End
   Begin VB.CommandButton cmdclsFileStream 
      Caption         =   "clsFileStream Example"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1620
      Width           =   3195
   End
   Begin VB.CommandButton CmdClsRegistryExample 
      Caption         =   "clsRegistry Example"
      Height          =   495
      Left            =   3660
      TabIndex        =   3
      Top             =   900
      Width           =   2955
   End
   Begin VB.CommandButton cmdClsFileSystem 
      Caption         =   "clsFileSystem Example"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   3195
   End
   Begin VB.CommandButton cmdClsIniExample 
      Caption         =   "clsIni Example"
      Height          =   495
      Left            =   3660
      TabIndex        =   1
      Top             =   180
      Width           =   2955
   End
   Begin VB.CommandButton cmdClsStrings 
      Caption         =   "clsStrings Example"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsFso As New clsFileSystem

Private Sub cmdClsCmdLine_Click()
    Dim clsCmd As New clsCmdLine
    Dim args '<-- this variant will hold an array
    
    Const cmdLine = "/bat -ball fred ""hi there"" john"
    'this command line has 8 arguments
    '/bat -> arguments preceeded by / are taken up to first space
    '-ball -> 4 arguments, those preceeded by - are taken as individual args
    'fred -> 1 argument, single words are taken as single arg up to space
    '"hi there" -> quoted strings are taken as single arguments
    'john -> final argument
    
    'On Startup it will automatically contain the programs
    'command line, here we explicitly set it with our commandline
    'above for example. You can also use this to allow users to enter
    'commands into your program at runtime.
    clsCmd.CommandLine = cmdLine
    
    'method 1 of use to get all the arguments into a variant array
    args = clsCmd.GetArgumentsToArray
    
    MsgBox "All Arguments: " & vbCrLf & vbCrLf & _
            Join(args, vbCrLf)
            
    'method 2 of use is to just test to see if an option is present
    MsgBox "/bat Option Chosen ? " & clsCmd.IsArgPresent("/bat")
    
    
End Sub

Private Sub cmdClsCmnDlg_Click()
    Dim cmndlg As New clsCmnDlg
    
    Dim openFile As String
    
    openFile = cmndlg.OpenDialog(textFiles, , "Open File", Me.hWnd)
    If Len(openFile) = 0 Then
        MsgBox "User Pressed Cancel"
    Else
        MsgBox openFile & " was selected"
    End If
    
    cmndlg.ErrorOnCancel = True
    On Error GoTo hadErr
    
    cmndlg.SetCustomFilter "Adobe PDF (*.pdf)", "*.pdf"
    openFile = cmndlg.OpenDialog(CustomFilter, App.Path)
    MsgBox openFile & " was selected"
    
    Exit Sub
hadErr: MsgBox "Error on cancel set to true and user select cancel"
End Sub

Private Sub cmdclsFileStream_Click()
    Dim fStream As New clsFileStream
    Dim clsFso As New clsFileSystem
    
    Dim exFile As String
    Dim i As Integer
    
    exFile = App.Path & "\example.ini"
    
    If Not clsFso.FileExists(exFile) Then
        MsgBox "Oops example file missing"
        Exit Sub
    End If
    
    With fStream
        .fOpen exFile, otreading
        
        While Not .EndOfFile 'read whole file line by line
            MsgBox .ReadLine
        Wend
         
        .fClose
        .fOpen exFile, otbinary
        
        While i < 5 'read first 5 characters
            MsgBox "Reading Character (" & i & ") = " & Chr(.BinGetChar)
            i = i + 1
        Wend
        .fClose
     
        Dim b() As String
        .fOpen exFile, otbinary
        
        ReDim b(.LengthOfFile)
        
        .BinGetStrArray b()
        
        'now b contains one character per array element
        For i = 0 To 4
            MsgBox b(i)
        Next
        
        .fClose
        
        
     End With
        
End Sub

Private Sub cmdClsFileSystem_Click()

    Dim pFolder, baseName, ext, fText
    Dim exFile As String, msg()
    
    exFile = App.Path & "\example.ini"
    
    With clsFso
        If Not .FileExists(exFile) Then
            MsgBox "Example File does not exist oops"
            Exit Sub
        End If
        
        pFolder = .GetParentFolder(exFile)
        baseName = .GetBaseName(exFile)
        ext = .GetExtension(exFile)
        fText = .readFile(exFile)
    End With
    
    push msg(), "Parent folder name is: " & pFolder
    push msg(), "File Base Name is: " & baseName
    push msg(), "File Extension is: " & ext
    push msg(), "File Text is: " & fText
    
    MsgBox Join(msg, vbCrLf)
        
End Sub

Private Sub cmdClsIniExample_Click()
    Dim clsIni As New clsIniFile
    Dim keys() As String
    
    Const sectName = "testSection"
    
    With clsIni
    
        .LoadFile App.Path & "\example.ini"
        keys() = .EnumKeys(sectName)
        
        MsgBox Join(keys, vbCrLf), , "Testsection Keys"
        
        MsgBox .GetValue(sectName, "key1"), , "Sect1 Key1="
    
        .Release
    
    End With

    'with this library you can create ini files on the fly, test to
    'see if keys or sections exist, and do just about anything else you
    'can think of with them.
    
    'Note that the enum functions return a 1 based array
    '
    'Note2: Just for safety sake, I would not use this on windows system
    '       ini files, I have used this on my own ini's without problem
    '       but it does store the whole thing in memory and dump the whole
    '       file to disk when done, if there was an err or system glitch
    '       in that process, the file could come out blank.
    '       This note only applies to writing changes to file
    '
    'Also note that all these expanded capabilities are only possible because
    '   it loads the entire file into memory as a block. This means that very large
    '   ini files will not be handled efficiently. I would not use this library on
    '   ini files over 200k (which is quite large!)
End Sub

Private Sub CmdClsRegistryExample_Click()
    Dim clsReg As New clsRegistry
    Dim keys() As String
    
    keys() = clsReg.EnumKeys(HKEY_LOCAL_MACHINE, "\Software")
    
    MsgBox "Keys under HKLM\Software: " & vbCrLf & vbCrLf & _
            Join(keys(), vbCrLf)
    
    If Not clsReg.KeyExists(HKEY_LOCAL_MACHINE, "\Software\Myapp") Then
        clsReg.CreateKey HKEY_LOCAL_MACHINE, "\Software\Myapp"
    End If
    
    clsReg.SetValue HKEY_LOCAL_MACHINE, "\Software\Myapp", "myKeyName", "myValue", REG_SZ
    
    MsgBox clsReg.ReadValue(HKEY_LOCAL_MACHINE, "\Software\Myapp", "myKeyName")
    
    clsReg.DeleteValue HKEY_LOCAL_MACHINE, "\Software\Myapp", "myvalue"
    clsReg.DeleteKey HKEY_LOCAL_MACHINE, "\Software\Myapp"
    
End Sub

Private Sub cmdClsStrings_Click()
    Dim c As clsStrings
    Const parseMe As String = "this is 'my string' and it needs to be parsed"
    
    Dim sqStr 'single quoted string
    Dim cnt   'count of how many ' are in teh parse string
    
    Set c = New clsStrings
    
    'set the string to parse
    c.Strng = parseMe
    
    'this function returns the index of the first
    'instance of the string we are looking for.
    'here, we are just using it to set the pointer
    'character (the character to look for) and the
    'pointer index (where a match was made)
    c.IndexOf "'"
    
    
    'the pointer char was set to ' by the above function.
    'the class also knows where the last match was made
    'it will now search the string for the next '
    sqStr = c.SubstringToNext
    MsgBox "Single Quoted String= " & sqStr
    
    cnt = c.CountOccurancesOf("'")
    MsgBox "# of single Quotes in string: " & cnt
    
    'there are many other functions in this library,
    'experiment away!
    
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub cmdHelp_Click()
    
    Dim hlpFile As String
    
    hlpFile = App.Path & "\..\vbDevKit.chm"
    
    If Not Len(Dir(hlpFile)) > 0 Then
        MsgBox "Could Not Locate HelpFile ", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "hh " & hlpFile, vbNormalFocus
    
    If Err.Number > 0 Then
        MsgBox "Error Starting helpFile:" & vbCrLf & vbCrLf & _
                Err.Description, vbInformation
    End If
    
End Sub
