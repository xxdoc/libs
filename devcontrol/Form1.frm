VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin devControl.ctlDevControl dev 
      Height          =   5730
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   10107
      intellisense    =   0   'False
      linenos         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'todo autoindent
'onscroll update line numbers..

Private Sub Form_Load()

    With dev
        .useIntellisense = False
        .useLineNumbers = True
        .isVbs = False
        
    End With
    
    
End Sub
