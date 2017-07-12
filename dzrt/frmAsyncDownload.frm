VERSION 5.00
Begin VB.Form frmAsyncDownload 
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   ScaleHeight     =   675
   ScaleWidth      =   2025
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox dl 
      Height          =   435
      Left            =   180
      ScaleHeight     =   375
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmAsyncDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'the usercontrol has the native async download methods..we must host it
'on a hidden form. then we expose the functionality via a class.

'i have decided to not expose the usercontrol publically so I could keep this project
'as a dll instead of ocx. OCXs seem to have stricter versioning and can lead to load
'/launch failures of compiled executables especially when binary compatability has not
'yet been set.

'this form is so that it can be used from CAsyncDownload.cls without requiring
'a user form

'this form is never shown and has not code since the usercontrol events are hooked
'in the class file..

