VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "lifeTest2"
      Height          =   375
      Left            =   6255
      TabIndex        =   8
      Top             =   4365
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "clear list"
      Height          =   330
      Left            =   6300
      TabIndex        =   7
      Top             =   3645
      Width           =   1185
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   225
      TabIndex        =   6
      Top             =   3015
      Width           =   5955
   End
   Begin VB.CommandButton Command4 
      Caption         =   "lifetime Test"
      Height          =   375
      Left            =   6300
      TabIndex        =   5
      Top             =   3105
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "mem leak test"
      Height          =   330
      Index           =   1
      Left            =   1530
      TabIndex        =   4
      Top             =   2565
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "native"
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   2565
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "mem leak test"
      Height          =   420
      Index           =   0
      Left            =   6300
      TabIndex        =   2
      Top             =   990
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   2265
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   180
      Width           =   5910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   6255
      TabIndex        =   0
      Top             =   270
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'after running this a bunch and watching memory usage
'and then attaching in a debugger and scanning memory for my test strings
'it appears we have proper ownership of the collections and its contents
'and its all deallocating automatically as it should and without any corruption
'due to dangling references from unexpected frees :)

'we cant create a vba.collection in C, so we use a reference passed in from vb
Private Declare Sub addItems Lib "col_dll" (ByRef col As Collection)

Dim hLib As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'watch the memory size in task manager while you run the tests to see if its leaking
Dim gCol As New Collection

Private Sub Command1_Click()

    Dim c As New Collection
    Dim x, tmp
    
    addItems c
    Me.Caption = c.Count & " items returned"
    
    For Each x In c
        tmp = tmp & x & vbCrLf
    Next
    
    Text1 = tmp
    
End Sub



Private Sub Command2_Click(index As Integer)
    
    For i = 0 To 100000
        If index = 0 Then
            Command1_Click
        Else
            Command3_Click
        End If
        DoEvents
        If i Mod 10 = 0 Then Command2(index).Caption = i
    Next
    
    Me.Caption = "complete!"
    
End Sub

Private Sub Command3_Click()
    
    Dim c As New Collection
    Dim x, tmp
    
    c.Add "native string 1"
    c.Add "native string 2"
    c.Add "native string 3"
    c.Add "native string 4"
    c.Add "native string 5"
    
    Me.Caption = c.Count & " items returned"
    
    For Each x In c
        tmp = tmp & x & vbCrLf
    Next
    
    Text1 = tmp
End Sub

Private Sub Command4_Click()
    
    Dim c() As Collection
    Dim x
    
    Const testSize = 10000
    
    ReDim c(testSize)
    
    For i = 0 To testSize
        Set c(i) = New Collection
        addItems c(i)
    Next
    
    List1.Clear
    For i = 0 To testSize
        For Each x In c(i)
            List1.AddItem CStr(x)
        Next
    Next
    
    MsgBox "Made it to end without crash!"
    
    'For i = 0 To testSize
    '    Set c(i) = Nothing
    'Next
    
    Erase c

    
End Sub

Private Sub Command5_Click()
    List1.Clear
    Set gCol = New Collection
End Sub

Private Sub Command6_Click()

    addItems gCol
    
    List1.Clear
    For Each x In gCol
       List1.AddItem x
    Next
    
    Me.Caption = gCol.Count
    
End Sub

'this bit below is for the IDEs benifit only. We make sure it finds out dll
'and also that it releases its interest in it when our ide debugging session
'ends. this allows us to recompile the C dll without having to shutdown the IDE
'each time..

Private Sub Form_Load()
    hLib = LoadLibrary("col_dll.dll")
    If hLib = 0 Then
        Me.Caption = "Failed to load dll"
        Command1.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hLib <> 0 Then FreeLibrary hLib
End Sub
