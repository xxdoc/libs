VERSION 5.00
Begin VB.UserControl ucProgress 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   810
   ScaleWidth      =   4800
   Begin VB.Shape s 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00008000&
      FillStyle       =   3  'Vertical Line
      Height          =   195
      Left            =   0
      Top             =   0
      Width           =   4155
   End
   Begin VB.Label lblPercent 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "lblPercent"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "ucProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright David Zimmer 2005
'site: http://sandsprite.com
'license: free for any use

'drop in replacement for mscomctl progressbar with some extras

Option Explicit
Private m_caption As String
Private m_max As Long
Private m_value As Long
Private m_showPercent As Boolean
Private lastRefresh As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Property Get caption() As String
    caption = m_caption
End Property

Property Let caption(text As String)
    m_caption = text
    m_showPercent = False
    If Len(text) = 0 Then
        lblPercent.Visible = False
    Else
        If UserControl.Height >= lblPercent.Height Then
            lblPercent.caption = text
            lblPercent.Left = (UserControl.Width / 2) - (lblPercent.Width / 2)
             lblPercent.Visible = True
        End If
    End If
End Property

Private Sub UserControl_Initialize()
      s.Visible = False
      s.Width = 0
      s.Height = UserControl.Height
      lblPercent.Visible = False
      lblPercent.Top = 0
End Sub

Private Sub UserControl_Resize()
   If Not Ambient.UserMode Then 'hosted on form in user IDE, not runtime
        s.Width = UserControl.Width
        s.Visible = True
   End If
   lblPercent.caption = "   "
   lblPercent.Left = (UserControl.Width / 2) - (lblPercent.Width / 2)
   lblPercent.Top = (UserControl.Height / 2) - (lblPercent.Height / 2) - 25
End Sub

Sub reset()
    m_value = 0
    s.Width = 0
    s.Visible = False
    lblPercent.Visible = False
End Sub

Property Get Max() As Long
    Max = m_max
End Property

Property Let Max(v As Long)
    If v < 0 Then v = 0
    Call reset
    m_max = v
End Property

Property Get Value() As Long
    Value = m_value
End Property

Property Let Value(v As Long)
    
    On Error Resume Next
    Dim maxWidth As Long
    Dim curWidth As Long
    Dim t As Long
    
    If v < 0 Then v = 0
    If v > m_max Then m_value = m_max Else m_value = v
    
    If v = 0 Then
        reset
        Exit Property
    End If
    
    If Not s.Visible Then
        s.Visible = True
        If m_showPercent And UserControl.Height >= lblPercent.Height Then lblPercent.Visible = True
    End If
    
    If v = m_max Then
        s.Width = UserControl.Width
    Else
        maxWidth = UserControl.Width
        curWidth = (m_value * maxWidth) / Max
        s.Width = curWidth
    End If
    
    If m_showPercent Then updatePercentage
    
    t = GetTickCount
    If t - lastRefresh > 150 Then 'eliminate some flicker use less cpu in tight loops
        UserControl.Refresh
        If m_showPercent Then lblPercent.Refresh
        DoEvents
    End If
    
    lastRefresh = t
    
End Property




'the rest of this file is extras and fluff stuff
'-------------------------------------------------------
Property Get ShowPercent() As Boolean
    ShowPercent = m_showPercent
End Property

Property Let ShowPercent(v As Boolean)
    m_showPercent = v
    If Not v Then
        lblPercent.Visible = False
    Else
        If UserControl.Height >= lblPercent.Height Then
            lblPercent.Visible = True
            updatePercentage
        End If
    End If
End Property

Sub updatePercentage()
    On Error Resume Next
    Dim pcent As Long
    pcent = (m_value / m_max) * 100
    lblPercent = pcent & "%"
End Sub

Sub inc(Optional ticks As Long = 1)
    Value = Value + ticks
End Sub

Sub dec(Optional ticks As Long = 1)
    Value = Value - ticks
End Sub

Sub setPercent(precentage As Long)
    On Error Resume Next
    
    If precentage <= 0 Then
        Value = 0
        Exit Sub
    End If
        
    If precentage >= 100 Then
        Value = m_max
        Exit Sub
    End If
    
    Dim v As Long
    v = (precentage * m_max) / 100
    Value = v
    
End Sub

Sub AssumeMinHeight()
    UserControl.Height = lblPercent.Height
End Sub

Property Get MinHeightForFontSize(fontSize As Long) As Long
    Dim tmp As Long
    tmp = lblPercent.fontSize
    lblPercent.fontSize = fontSize
    MinHeightForFontSize = lblPercent.fontSize
    lblPercent.fontSize = tmp
End Property

Property Let fontSize(v As Long)
    On Error Resume Next
    lblPercent.fontSize = v
    UserControl_Resize
End Property

Property Get fontSize() As Long
    fontSize = lblPercent.fontSize
End Property
    
Property Get FillStyle() As Long
    FillStyle = s.FillStyle
End Property

Property Let FillStyle(v As Long)
    On Error Resume Next
    If v < 0 Or v > 7 Then v = 3 'default to vertical if invalid..
    s.FillStyle = v
End Property

Property Get FillColor() As Long
    FillColor = s.FillColor
End Property

Property Let FillColor(v As Long)
    On Error Resume Next
    s.FillColor = v
End Property

Property Get BackColor() As Long
     BackColor = s.BackColor
End Property

Property Let BackColor(v As Long)
    On Error Resume Next
    s.BackColor = v
End Property

