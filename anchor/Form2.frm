VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   8085
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   1995
      Left            =   2610
      TabIndex        =   5
      Text            =   "This one has no tag"
      Top             =   2790
      Width           =   5370
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6885
      TabIndex        =   4
      Tag             =   "0,1,1"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5670
      TabIndex        =   3
      Tag             =   "0,1,1"
      Top             =   7380
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   2445
      Left            =   2610
      TabIndex        =   2
      Tag             =   "1,1"
      Text            =   "this one has only first 2  attributes set in tag"
      Top             =   180
      Width           =   5370
   End
   Begin VB.TextBox Text1 
      Height          =   2220
      Left            =   2610
      TabIndex        =   1
      Tag             =   "0,1,0"
      Text            =   "Text1"
      Top             =   4995
      Width           =   5370
   End
   Begin VB.ListBox List1 
      Height          =   7080
      Left            =   45
      TabIndex        =   0
      Tag             =   "1,0,1"
      Top             =   135
      Width           =   2310
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim anchor As New CAnchor

Private Sub Form_Load()
        
    'this example uses resize attributes saved in
    'each form elements .tag, partial or no attributes are ok
    'as long as you understand the default behaviors
    
    anchor.LoadSettingsFromTags Me
    
End Sub

