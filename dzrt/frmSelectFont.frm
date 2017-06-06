VERSION 5.00
Begin VB.Form frmSelectFont 
   Caption         =   "Select Font"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   180
      ScaleHeight     =   795
      ScaleWidth      =   3015
      TabIndex        =   15
      Top             =   2820
      Width           =   3075
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sample Text"
         Height          =   195
         Left            =   1020
         TabIndex        =   16
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      Height          =   855
      Left            =   3960
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Appearance"
      Height          =   1695
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox Check4 
         Caption         =   "Strikethru"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Font"
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtFont 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Font Size"
      Height          =   2535
      Left            =   2400
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmSelectFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' based on a psc submission by www.computing.iscute.com

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

'constants for searching the ListBox
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const LB_FINDSTRING = &H18F


Dim loaded As Boolean
Dim abort As Boolean
Dim dlg As CCmnDlg

Private Sub Check1_Click()
    lblFont.FontBold = IIf(Check1.value = 1, True, False)
    center
End Sub

Private Sub Check2_Click()
    lblFont.FontItalic = IIf(Check2.value = 1, True, False)
    center
End Sub

Private Sub Check3_Click()
    lblFont.FontUnderline = IIf(Check3.value = 1, True, False)
    center
End Sub

Private Sub Check4_Click()
    lblFont.FontStrikethru = IIf(Check4.value = 1, True, False)
    center
End Sub

Function SelectFont(cmndlg As CCmnDlg, Optional obj As Object) As CFont
    On Error Resume Next
    
    Dim isBold As Boolean
    Dim isItalic As Boolean
    Dim isUnderLine As Boolean
    Dim isStrike As Boolean
    Dim color As Long
    Dim name As String
    Dim size As Long
    Dim f As New CFont
    
    Set dlg = cmndlg
    Set SelectFont = f
    If Not loaded Then Form_Load
    
    If Not obj Is Nothing Then
        isBold = obj.FontBold
        If Err.Number <> 0 Then
            isBold = obj.Font.Bold
            Err.Clear
        End If
        If isBold Then Check1.value = 1
        
        isItalic = obj.FontItalic
        If Err.Number <> 0 Then
            isItalic = obj.Font.Italic
            Err.Clear
        End If
        If isItalic Then Check2.value = 1
        
        isUnderLine = obj.FontUnderline
        If Err.Number <> 0 Then
            isUnderLine = obj.Font.Underline
            Err.Clear
        End If
        If isUnderLine Then Check3.value = 1
    
        isStrike = obj.FontStrikethru
        If Err.Number <> 0 Then
            isStrike = obj.Font.Strikethru
            Err.Clear
        End If
        If isStrike Then Check4.value = 1
    
        color = obj.ForeColor
        If Err.Number <> 0 Then
            color = obj.Font.ForeColor
            Err.Clear
        End If
        
        name = obj.fontname
        If Err.Number <> 0 Then
            name = obj.Font.name
            Err.Clear
        End If
        
        size = obj.FontSize
        If Err.Number <> 0 Then
            size = obj.Font.size
            Err.Clear
        End If
    
        If Len(name) > 0 Then txtFont = name
        If size <> 0 Then Text1 = size
        If color <> 0 Then
            Label3.BackColor = obj.ForeColor
            lblFont.ForeColor = obj.ForeColor
        End If
        
        For x = 0 To List2.ListCount - 1
            If size = val(List2.List(x)) Then
                List2.ListIndex = x
                lblFont.FontSize = val(List2.List(x))
                Text1.Text = List2.List(x)
                Exit For
            End If
        Next
    
    End If
    
    abort = False
    Me.Show 1
    
    If abort Then
        Unload Me
        Exit Function
    End If
    
    With f
        .Bold = lblFont.FontBold
        .Italic = lblFont.FontItalic
        .Underline = lblFont.FontUnderline
        .Strikethrough = lblFont.FontStrikethru
        .name = lblFont.fontname
        .size = lblFont.FontSize
        .color = lblFont.ForeColor
    End With
    
    Unload Me
    
End Function

Private Sub Command1_Click()
    Me.Visible = False
End Sub

Private Sub Command2_Click()
    abort = True
    Unload Me
End Sub

Private Sub Form_Load()

    If List1.ListCount > 0 Then Exit Sub
      
    For x = 0 To Screen.FontCount - 1
       List1.AddItem Screen.Fonts(x)
    Next
               
End Sub

Private Sub Form_Unload(Cancel As Integer)
    abort = True
End Sub

Private Sub Label3_Click()
    Dim h As Long
    h = dlg.ColorDialog()
    If h <> 0 Then
        Label3.BackColor = h
        lblFont.ForeColor = h
    End If
End Sub

Private Sub List1_Click()
    On Error Resume Next
    Dim c As Collection
    List2.Clear
    lblFont.fontname = List1.List(List1.ListIndex)
    Set c = EnumFontSizes(lblFont.fontname)
    For Each x In c
        List2.AddItem x
    Next
    center
End Sub

Private Sub List2_Click()
    Text1.Text = List2.List(List2.ListIndex)
    lblFont.FontSize = val(Text1.Text)
    center
End Sub

Private Sub Text1_Change()
    
    On Error Resume Next
    
    For x = 0 To List2.ListCount - 1
        If val(Text1.Text) = val(List2.List(x)) Then
             List2.ListIndex = x
             
             Exit For
        End If
    Next
    
    lblFont.FontSize = val(Text1.Text)
    
End Sub

Private Sub txtFont_Change()
    On Error Resume Next
    List1.ListIndex = GetListBoxIndex(List1.hWnd, txtFont.Text)
End Sub

'function to get find an item in the Listbox
Public Function GetListBoxIndex(hWnd As Long, SearchKey As String, Optional FindExactMatch As Boolean = False) As Long

    If FindExactMatch Then
        GetListBoxIndex = SendMessage(hWnd, LB_FINDSTRINGEXACT, -1, ByVal SearchKey)
    Else
        GetListBoxIndex = SendMessage(hWnd, LB_FINDSTRING, -1, ByVal SearchKey)
    End If

End Function

Function center()
    On Error Resume Next
    Dim w As Long, h As Long, lw As Long, lh As Long
    With lblFont
        lh = .Height
        lw = .Width
        w = Picture1.Width
        h = Picture1.Height
        
        If h > .Height Then
            .Top = ((h - .Height) / 2) - 50
        Else
            .Top = 0
            Picture1.Height = lh
            Me.Height = Picture1.Top + Picture1.Height + 600
        End If
        
        If w > .Width Then
            .Left = (w - .Width) / 2
        Else
            .Left = 0
        End If
    End With
    
End Function
