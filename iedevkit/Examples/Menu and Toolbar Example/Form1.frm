VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tool Bar Button and Context Menu demos"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2235
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Form1.frx":0000
      Top             =   2700
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   2355
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "Form1.frx":01CA
      Top             =   180
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Context Menu Demo"
      Height          =   2235
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1875
      Begin VB.CommandButton CmdEnumerateContextMenus 
         Caption         =   "Enumerate"
         Height          =   495
         Left            =   300
         TabIndex        =   7
         Top             =   1560
         Width           =   1275
      End
      Begin VB.CommandButton cmdRemoveContextMenu 
         Caption         =   "Remove"
         Height          =   495
         Left            =   300
         TabIndex        =   6
         Top             =   900
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddToContextMenu 
         Caption         =   "Add"
         Height          =   495
         Left            =   300
         TabIndex        =   5
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tool Bar Button Demo"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1875
      Begin VB.CommandButton cmdAddToolBarButton 
         Caption         =   "Add"
         Height          =   495
         Left            =   300
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdModifyToolBarButton 
         Caption         =   "Modify"
         Height          =   495
         Left            =   300
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemoveToolBarButton 
         Caption         =   "Remove"
         Height          =   495
         Left            =   300
         TabIndex        =   1
         Top             =   1620
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsIntegrate As New clsIeIntegration

Dim myGuid As String 'key to which toolbar button we added dont loose
 
Private Sub cmdAddToolBarButton_Click()
    
     myGuid = clsIntegrate.IntegrateWithIEToolbar( _
                "Example Toolbar", _
                App.Path & "\hot.ico", _
                App.Path & "\cold.ico", _
                App.Path & "\example.html", True)

End Sub

Private Sub cmdModifyToolBarButton_Click()
    clsIntegrate.ModifyToolBarSetting myGuid, "Modified Example Toolbar"
End Sub

Private Sub cmdRemoveToolBarButton_Click()
    clsIntegrate.RemoveIEToolBar myGuid
End Sub

Private Sub cmdAddToContextMenu_Click()
     
     clsIntegrate.AddToRightClickMenu "Always there", _
                App.Path & "\example.html", swDefault
                
     clsIntegrate.AddToRightClickMenu "Only For Links", _
                App.Path & "\exampleUI.html", swAnchor, True
                
End Sub

Private Sub cmdRemoveContextMenu_Click()
    With clsIntegrate
        .RemoveFromRightClickMenu "Always there"
        .RemoveFromRightClickMenu "Only For Links"
    End With
End Sub

Private Sub CmdEnumerateContextMenus_Click()
    Dim ret() As String
    ret() = clsIntegrate.ShowInstalledMenuExtensions()
    MsgBox Join(ret, vbCrLf)
End Sub

