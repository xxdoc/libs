VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIntellisense 
   BorderStyle     =   0  'None
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrLostFocus 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   2940
      Top             =   2220
   End
   Begin VB.Frame fraLv 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.Frame coverup 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -360
         TabIndex        =   1
         Top             =   1440
         Width           =   3015
      End
      Begin MSComctlLib.ListView lv 
         Height          =   1875
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   3307
         View            =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmIntellisense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'copyright David Zimmer <dzzie@yahoo.com> 2001

'THIS IS USED IN MULTIPLE PROJECTS SO EDIT CAREFULLY !

Private selStart As Long
Private typedText As String
Private pDevControl As ctlDevControl

Public isVisible As Boolean
'Public isMDIChild As Boolean
'Public MDI_OffsetLeft As Long
'Public MDI_OffsetTop As Long

Private Sub ShowIt()
    
    Me.Visible = True
    SetWindowTopMost Me
    tmrLostFocus.Enabled = True
    isVisible = True
    
End Sub

Private Sub HideIt()
    Me.Visible = False
    isVisible = False
End Sub

Sub Display()
    ShowIt
End Sub

Private Sub lv_DblClick()
    lv_KeyPress 13
End Sub

Private Sub lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'stops listview items from being dragged about
    PostMessage lv.hwnd, WM_LBUTTONUP, 0&, 0&
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
    
'    MsgBox KeyAscii
    
    Dim rtf As RichTextBox
    
    If KeyAscii = 8 Or KeyAscii = 27 Then 'delete or esc key just hide
        'because used in multiple projects no links to frmMain
        'modSyntaxHighlighting.pDevControl is part of framework and
        'will always be there
        Set rtf = pDevControl.clsRtf.GetRtf
        With rtf
            .selStart = selStart + Len(typedText)
            .SetFocus
        End With
        Unload Me
    ElseIf KeyAscii = 13 Or KeyAscii = 32 Then
        HideIt
        pDevControl.Intellisense_SelectionMade lv.SelectedItem.Text, lv.SelectedItem.Tag
        Unload Me
    Else
        Set rtf = pDevControl.clsRtf.GetRtf
        With rtf
            typedText = typedText & Chr(KeyAscii)
            .selText = typedText
            .selStart = selStart
            .selLength = Len(typedText)
        End With
    End If
    
End Sub


Sub ResizeAndActivate(numElem As Integer, parentDevControl As ctlDevControl) 'intellisense window
   
    Dim sbHeight As Integer
    Dim txtLen As Integer
    
    Dim x, y
    
    Set pDevControl = parentDevControl
    pDevControl.GetCaretPos x, y

    sbHeight = ScrollBarHeightTwips()
     
    lv.Height = IIf(numElem < 10, numElem * 290, 2500)
        
    Dim offsetX, offsetY
    
    If pDevControl.MDI_OffsetLeft > 0 Then
    
        offsetX = pDevControl.MDI_OffsetLeft
        offsetY = pDevControl.MDI_OffsetTop + 1100  'menubar & toolbar
            
        'If config.showExplorer Then
        '    offsetX = offsetX + frmMain.picExplorerContainer.Width
        'End If
        
    End If
        
        
    Me.Move offsetX + x + 120, offsetY + y + 120
        
    fraLv.Width = lv.Width + 30
    Me.Width = fraLv.Width + 20
    Me.Display
    DoEvents
    
    If ListviewHasHScroll(lv) Then
        fraLv.Height = lv.Height - sbHeight + 30
        coverup.top = fraLv.Height - 30
        Height = fraLv.Height
    Else
        coverup.top = fraLv.Height + 375
        fraLv.Height = lv.Height
        Me.Height = fraLv.Height
    End If
        
    lv.ListItems(1).Selected = True
    lv.ListItems(1).EnsureVisible
    lv.SetFocus
    
    selStart = pDevControl.clsRtf.GetRtf.selStart
    typedText = Empty
    
End Sub

'Private Function LongestLVText() As Long
'    Dim i, ret, tmp
'    For i = 1 To lv.ListItems.Count
'        tmp = Len(lv.ListItems(i).text)
'        If tmp > ret Then ret = tmp
'    Next
'    LongestLVText = ret
'End Function
'
'If uMsg = WM_ACTIVATEAPP Then
'             'Check to see if Activating the application
'             If wParam <> 0 Then
'                 'Application Received Focus
'                 Form1.Caption = "Focus Restored"
'             Else
'                 'Application Lost Focus
'                 Form1.Caption = "Focus Lost"
'             End If
'         End If

Private Sub tmrLostFocus_Timer()
    If GetForegroundWindow() <> Me.hwnd Then Unload Me
End Sub
