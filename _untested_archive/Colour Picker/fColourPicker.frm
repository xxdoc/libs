VERSION 5.00
Begin VB.Form fColourPicker 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Picker"
   ClientHeight    =   5235
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   ShowInTaskbar   =   0   'False
   Begin vhColourPicker.ucColourPicker ucColourPicker 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6855
      _extentx        =   12091
      _extenty        =   9340
   End
End
Attribute VB_Name = "fColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cancelled As Boolean
Private Sub Form_Paint()
   ucColourPicker.Refresh
End Sub
Private Sub ucColourPicker_DialogClosed(UserCancelled As Boolean)
   Me.Hide
   Cancelled = UserCancelled
End Sub
