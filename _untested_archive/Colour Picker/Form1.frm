VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3915
      Left            =   240
      ScaleHeight     =   3885
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Click to set PictureBox colour"
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   4260
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColourPicker As cColourPicker
Private Sub Form_Load()
   Set ColourPicker = New cColourPicker
End Sub
Private Sub Form_Terminate()
  If Forms.Count = 0 Then New_c.CleanupRichClientDll
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set ColourPicker = Nothing
End Sub
Private Sub Picture1_Click()
Dim ret As Long
   ret = ColourPicker.Show(Picture1.BackColor)
   If Not ColourPicker.Cancelled Then
      Picture1.BackColor = ret
   End If
End Sub
