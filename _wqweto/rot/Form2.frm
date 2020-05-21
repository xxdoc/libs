VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   510
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   495
      TabIndex        =   0
      Top             =   360
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
     
    Private m_oForm1 As Object
     
    Private Sub Form_Load()
        Set m_oForm1 = GetObject("MySpecialProject.Form1")
    End Sub
     
    Private Sub Command1_Click()
        m_oForm1.List1.AddItem "Test " & Now
    End Sub
     
    Private Sub Command2_Click()
        m_oForm1.List1.RemoveItem 0
    End Sub
