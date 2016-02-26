VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6885
      TabIndex        =   4
      Top             =   5670
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Form2 Example"
      Height          =   375
      Left            =   5310
      TabIndex        =   3
      Top             =   5670
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   3075
      Left            =   2565
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   135
      Width           =   5370
   End
   Begin VB.TextBox Text1 
      Height          =   2220
      Left            =   2565
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3375
      Width           =   5325
   End
   Begin VB.ListBox List1 
      Height          =   5910
      Left            =   45
      TabIndex        =   0
      Tag             =   "tes tag"
      Top             =   135
      Width           =   2310
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim anchor As New CAnchor

Private Sub Command1_Click()
    Form2.Visible = True
End Sub

Private Sub Form_Load()
        
    'you can actually reduce this down to one line of code
    'per form, see the form2 example. .tag attributes are saved in the frm file
    'note this is very subtle though, unfamiliar devs would have no idea wtf was going on
    
    Set anchor.owner = Me
    anchor.AddItem List1, True, , True
    anchor.AddItem Text2
    anchor.AddItem Text1, False, True
    anchor.AddItem Command1, False, True, True
    anchor.AddItem Command2, False, True, True
    
End Sub

