VERSION 5.00
Begin VB.Form frmDynamicMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dynamic Menu"
   ClientHeight    =   1545
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3840
   Icon            =   "DynamicMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDumpToDebug 
      Caption         =   "Dump  to Debug Window"
      Height          =   375
      Left            =   803
      TabIndex        =   2
      Top             =   720
      Width           =   2235
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Test"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddMenu 
      Caption         =   "Add Item to Menu Bar"
      Height          =   375
      Left            =   803
      TabIndex        =   0
      Top             =   240
      Width           =   2235
   End
   Begin VB.Menu mnuSubMenu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuMenuItem 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "frmDynamicMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Using the Menu APIs to Grow or Shrink a Menu During Run-time
'(c) Jon Vote, 2003
'
'Idioma Software Inc.
'jon@idioma-software.com
'www.idioma-software.com
'www.skycoder.com

Option Explicit

'Subclass the form
Private Sub Form_Load()

  Dim lngRC As Long
  
  'Make the form auto refresh
  Me.AutoRedraw = True
  
  'Subclass the form - this will
  'cause all Windows messages to be
  'trapped and processed by the
  'SubClassHandler routine.
  lngRC = SubClass(Me)
  
End Sub

'Add a sub-menu to the menu bar
Private Sub cmdAddMenu_Click()

  Dim strSubMenuCaption As String
  Dim strMenuItemCaption As String
    
  'Prompt the user for the sub-menu/menu item captions
  If GetSubMenuCaptions(strSubMenuCaption, strMenuItemCaption) Then
    'First one is already there - just make it visible
    If Not mnuSubMenu.Visible Then
      mnuSubMenu.Visible = True
      mnuSubMenu.Caption = strSubMenuCaption
      mnuMenuItem.Caption = strMenuItemCaption
    Else
      'Use the APIs to create and add a new sub-menu
      AddPopupMenu GetMenu(Me.hWnd), strSubMenuCaption, strMenuItemCaption
    End If
  End If
  
  'This is needed to refresh the menu bar
  DrawMenuBar Me.hWnd
  
End Sub

'Dump menu structure to debug window
Private Sub cmdDumpToDebug_Click()
  
  Debug.Print ""
  Debug.Print "Menu Caption", "Menu Handle", "Menu ID", "Position"
  Debug.Print "------------", "-----------", "-------", "--------"
  
  DumpMenu GetMenu(Me.hWnd)
  
End Sub

'Restore windows message handler
Private Sub Form_Unload(Cancel As Integer)

  Dim lngRC As Long
  
  lngRC = UnSubClass(Me)
  
End Sub
