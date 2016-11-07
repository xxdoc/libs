VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   510
      Left            =   3645
      TabIndex        =   8
      Top             =   1170
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "This Is My New Test String"
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "20000"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test STRINGEX"
      Height          =   480
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test standard STRING"
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "text to append"
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   300
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "iterations"
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   780
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   2070
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   1350
      Width           =   90
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'WINAPI

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

'EVENTS

Private Sub Command1_Click()

 

    Label1.Caption = TestString(CLng(Text1.Text))

End Sub

Private Sub Command2_Click()

    Label2.Caption = TestStringEx(CLng(Text1.Text))

End Sub

'ROUTINES

Private Function TestString(ByVal Iterations As Long) As String

    Dim i As Long
    Dim s As String
    Dim x As Long
    
    x = GetTickCount
    
    For i = 0& To Iterations
        
        s = s & Text2.Text
    
    Next i
    
    TestString = Len(s) & " characters in " & (GetTickCount - x) / 1000& & " seconds"

End Function

Private Function TestStringEx(ByVal Iterations As Long) As String

    Dim i As Long
    Dim s As StringEx
    Dim x As Long
    
    Set s = New StringEx
    x = GetTickCount
    
    For i = 0& To Iterations
        
        s.Concat Text2.Text
    
    Next i
    
    TestStringEx = s.length & " characters in " & (GetTickCount - x) / 1000& & " seconds"
    
    Set s = Nothing

End Function

Private Sub Command3_Click()
    Dim s As New StringEx
    
    s = "fart knocker"
    'Debug.Print s.subString(2)
    'Debug.Print s.subString(2, 6)
    'Debug.Print s.subString(2, 200)
    
    Dim b() As Byte
    b = s.ToArray(False)
    Debug.Print HexDump(b)
    
    
    
End Sub
