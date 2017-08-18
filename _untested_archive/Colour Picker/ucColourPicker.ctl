VERSION 5.00
Begin VB.UserControl ucColourPicker 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   4140
      TabIndex        =   45
      Top             =   60
      Width           =   855
      Begin VB.TextBox txtB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   300
         TabIndex        =   12
         Top             =   1020
         Width           =   480
      End
      Begin VB.TextBox txtG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   300
         TabIndex        =   10
         Top             =   540
         Width           =   480
      End
      Begin VB.TextBox txtR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "&R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   105
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "&G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   585
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "&B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   11
         Top             =   1065
         Width           =   120
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   11
         Left            =   60
         TabIndex        =   48
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   47
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   8
         Left            =   60
         TabIndex        =   46
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   540
      TabIndex        =   42
      Top             =   60
      Width           =   3255
      Begin VB.TextBox txtS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Top             =   60
         Width           =   480
      End
      Begin VB.TextBox txtL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   60
         Width           =   480
      End
      Begin VB.TextBox txtH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "&H"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   105
         Width           =   120
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   43
         Top             =   60
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "&S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   2220
         TabIndex        =   5
         Top             =   105
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "&L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   105
         Width           =   240
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   1200
         TabIndex        =   44
         Top             =   60
         Width           =   240
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   2160
         TabIndex        =   13
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Basic Colours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1635
      Left            =   4200
      TabIndex        =   26
      Top             =   1560
      Width           =   2415
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1440
         TabIndex        =   41
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1020
         TabIndex        =   40
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   39
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   38
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   600
         TabIndex        =   37
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   180
         TabIndex        =   36
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   1860
         TabIndex        =   35
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   1020
         TabIndex        =   34
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   1440
         TabIndex        =   33
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   1860
         TabIndex        =   32
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   31
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   600
         TabIndex        =   30
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   1020
         TabIndex        =   29
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   13
         Left            =   1440
         TabIndex        =   28
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lblSwatch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   1860
         TabIndex        =   27
         Top             =   1140
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   25
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5520
      TabIndex        =   24
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recent Picks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   4200
      TabIndex        =   0
      Top             =   3300
      Width           =   2415
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   1860
         TabIndex        =   23
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   180
         TabIndex        =   22
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   600
         TabIndex        =   21
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1440
         TabIndex        =   20
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   1020
         TabIndex        =   19
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   1440
         TabIndex        =   18
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   1860
         TabIndex        =   17
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblCustom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1020
         TabIndex        =   14
         Top             =   300
         Width           =   360
      End
   End
End
Attribute VB_Name = "ucColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'just a generic Helper-Control, to make Widgets usable on a normal VB-Form

Event DialogClosed(UserCancelled As Boolean)
Private Root As cWidgetRoot
 
Private Declare Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long
Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Long, ByVal wLuminance As Long, ByVal wSaturation As Long) As Long

Private mCancelled As Boolean
Private mBusy As Boolean

Private RecentColours As cCollection
Private WithEvents ColorChooser As cwColorChooser
Attribute ColorChooser.VB_VarHelpID = -1
Private WithEvents SSlider As cwSlider
Attribute SSlider.VB_VarHelpID = -1
Private WithEvents RSlider As cwSlider
Attribute RSlider.VB_VarHelpID = -1
Private WithEvents GSlider As cwSlider
Attribute GSlider.VB_VarHelpID = -1
Private WithEvents BSlider As cwSlider
Attribute BSlider.VB_VarHelpID = -1
'============================================================================
'PUBLIC Interface
Public Sub Refresh()
   ColorChooser.Colour = ColorChooser.Colour
   RSlider.Widget.Refresh
   GSlider.Widget.Refresh
   BSlider.Widget.Refresh
   SSlider.Widget.Refresh
End Sub
Public Property Get Widgets() As cWidgets
  If Root Is Nothing Then
    Set Root = Cairo.WidgetRoot
    Root.BackColor = vbWhite
    Root.RenderContentIn Me
  End If
  Set Widgets = Root.Widgets
End Property
Public Property Get ChosenColour() As Long
   ChosenColour = ColorChooser.Colour
End Property
Public Property Let InitialColour(pColour As Long)
Dim i As Long
   ColorChooser.Colour = pColour
   For i = 0 To RecentColours.Count - 1
      lblCustom(i).BackColor = RecentColours.ItemByIndex(i)
   Next i
End Property
'============================================================================
'COMMAND BUTTON Handling
Private Sub cmdCancel_Click()
   RaiseEvent DialogClosed(True)
End Sub
Private Sub cmdOK_Click()
   If Not RecentColours.Exists("Colour:" & ColorChooser.Colour) Then
      If RecentColours.Count = 10 Then RecentColours.RemoveByIndex 0
      RecentColours.Add ColorChooser.Colour, "Colour:" & ColorChooser.Colour
   End If
   RaiseEvent DialogClosed(False)
End Sub
'============================================================================
'TEXTBOX Handling
Private Sub txtH_GotFocus()
   txtH.SelStart = 0
   txtH.SelLength = Len(txtH)
End Sub
Private Sub txtH_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyUp Then
      ColorChooser.Hue = ColorChooser.Hue + 1
   ElseIf KeyCode = vbKeyDown Then
      ColorChooser.Hue = ColorChooser.Hue - 1
   Else
      Exit Sub
   End If
   KeyCode = 0
   txtH.SelStart = 0
   txtH.SelLength = Len(txtH)
End Sub
Private Sub txtH_KeyPress(KeyAscii As Integer)
   IgnoreNonNumericKeyPress KeyAscii
End Sub
Private Sub txtH_Validate(Cancel As Boolean)
   If Not IsNumeric(txtH) Then Exit Sub
   ColorChooser.Hue = txtH
End Sub
Private Sub txtL_GotFocus()
   txtL.SelStart = 0
   txtL.SelLength = Len(txtL)
End Sub
Private Sub txtL_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyUp Then
      ColorChooser.Luminance = ColorChooser.Lum + 1
   ElseIf KeyCode = vbKeyDown Then
      ColorChooser.Luminance = ColorChooser.Lum - 1
   Else
      Exit Sub
   End If
   KeyCode = 0
   txtL.SelStart = 0
   txtL.SelLength = Len(txtL)
End Sub
Private Sub txtL_KeyPress(KeyAscii As Integer)
   IgnoreNonNumericKeyPress KeyAscii
End Sub
Private Sub txtL_Validate(Cancel As Boolean)
   If Not IsNumeric(txtL) Then Exit Sub
   If txtL < 0 Then txtL = 0
   If txtL > 240 Then txtL = 240
   ColorChooser.Luminance = txtL
End Sub
Private Sub txtS_GotFocus()
   txtS.SelStart = 0
   txtS.SelLength = Len(txtS)
End Sub
Private Sub txtS_KeyDown(KeyCode As Integer, Shift As Integer)
   TextBoxKeyDownHandler txtS, KeyCode, SSlider
End Sub
Private Sub txtS_KeyPress(KeyAscii As Integer)
   IgnoreNonNumericKeyPress KeyAscii
End Sub
Private Sub txtS_Validate(Cancel As Boolean)
   TextBoxValidation txtS, SSlider, 0, 240
End Sub
Private Sub txtR_GotFocus()
   txtR.SelStart = 0
   txtR.SelLength = Len(txtR)
End Sub
Private Sub txtR_KeyDown(KeyCode As Integer, Shift As Integer)
   TextBoxKeyDownHandler txtR, KeyCode, RSlider
End Sub
Private Sub txtR_KeyPress(KeyAscii As Integer)
   IgnoreNonNumericKeyPress KeyAscii
End Sub
Private Sub txtR_Validate(Cancel As Boolean)
   TextBoxValidation txtR, RSlider, 0, 255
End Sub
Private Sub txtG_GotFocus()
   txtG.SelStart = 0
   txtG.SelLength = Len(txtG)
End Sub
Private Sub txtG_KeyDown(KeyCode As Integer, Shift As Integer)
   TextBoxKeyDownHandler txtG, KeyCode, GSlider
End Sub
Private Sub txtG_KeyPress(KeyAscii As Integer)
   IgnoreNonNumericKeyPress KeyAscii
End Sub
Private Sub txtG_Validate(Cancel As Boolean)
   TextBoxValidation txtG, GSlider, 0, 255
End Sub
Private Sub txtB_GotFocus()
   txtB.SelStart = 0
   txtB.SelLength = Len(txtB)
End Sub
Private Sub txtB_KeyDown(KeyCode As Integer, Shift As Integer)
   TextBoxKeyDownHandler txtB, KeyCode, BSlider
End Sub
Private Sub txtB_KeyPress(KeyAscii As Integer)
   IgnoreNonNumericKeyPress KeyAscii
End Sub
Private Sub txtB_Validate(Cancel As Boolean)
   TextBoxValidation txtB, BSlider, 0, 255
End Sub
Private Sub IgnoreNonNumericKeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then
      KeyAscii = 0
   End If
End Sub
Private Sub TextBoxKeyDownHandler(TxtBx As TextBox, KeyCode As Integer, Slider As cwSlider)
   If KeyCode = vbKeyUp Then
      Slider.Value = Slider.Value + 1
   ElseIf KeyCode = vbKeyDown Then
      Slider.Value = Slider.Value - 1
   Else
      Exit Sub
   End If
   KeyCode = 0
   TxtBx.SelStart = 0
   TxtBx.SelLength = Len(TxtBx)
End Sub
Private Sub TextBoxValidation(TxtBx As TextBox, Slider As cwSlider, MinValue As Long, MaxValue As Long)
   If Not IsNumeric(TxtBx) Then Exit Sub
   If TxtBx < MinValue Then TxtBx = MinValue
   If TxtBx > MaxValue Then TxtBx = MaxValue
   Slider.Value = TxtBx
End Sub
'============================================================================
'SLIDER Handling
Private Sub RSlider_ValueChanged(ByVal NewValue As Double)
   If mBusy Then Exit Sub
   ColorChooser.Red = CInt(NewValue)
End Sub
Private Sub GSlider_ValueChanged(ByVal NewValue As Double)
   If mBusy Then Exit Sub
   ColorChooser.Green = CInt(NewValue)
End Sub
Private Sub BSlider_ValueChanged(ByVal NewValue As Double)
   If mBusy Then Exit Sub
   ColorChooser.Blue = CInt(NewValue)
End Sub
Private Sub SSlider_ValueChanged(ByVal NewValue As Double)
   If mBusy Then Exit Sub
   ColorChooser.Saturation = CInt(NewValue) '+ 1
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Increment As Long
   Increment = IIf(Shift = 1, -1, 1)
   Select Case KeyCode
      Case vbKeyEscape
         cmdCancel_Click
      Case vbKeyReturn
         cmdOK_Click
      Case vbKeyR
         RSlider.Value = RSlider.Value + Increment
      Case vbKeyG
         GSlider.Value = GSlider.Value + Increment
      Case vbKeyB
         BSlider.Value = BSlider.Value + Increment
      Case vbKeyH
         ColorChooser.Hue = ColorChooser.Hue + Increment
      Case vbKeyL
         ColorChooser.Luminance = ColorChooser.Lum + Increment
      Case vbKeyS
         SSlider.Value = SSlider.Value + Increment
   End Select
End Sub
Private Sub ColorChooser_ColorChanged(PickedColour As Long)
   UpdateLabels PickedColour
End Sub
Private Sub lblCustom_Click(Index As Integer)
   ColorChooser.Colour = lblCustom(Index).BackColor
End Sub
Private Sub lblSwatch_Click(Index As Integer)
   ColorChooser.Colour = lblSwatch(Index).BackColor
End Sub
Private Sub UpdateLabels(pColour As Long)
Dim R As Double, G As Double, B As Double
   mBusy = True
   Cairo.ColorSplit pColour, R, G, B
   txtR.Text = CInt(R * 255)
   txtG.Text = CInt(G * 255)
   txtB.Text = CInt(B * 255)
   RSlider.Value = CInt(R * 255)
   GSlider.Value = CInt(G * 255)
   BSlider.Value = CInt(B * 255)
   
   txtH.Text = ColorChooser.Hue
   txtL.Text = ColorChooser.Lum
   txtS.Text = ColorChooser.Sat
   SSlider.Value = ColorChooser.Sat
   Root.Widget.Refresh
   mBusy = False
End Sub
'============================================================================
'STARTUP /TEAR-DOWN
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim H As Long
   If Not Ambient.UserMode Then Exit Sub
   
   For H = 0 To 220 Step 20
      lblSwatch(H / 20).BackColor = ColorHLSToRGB(H, 120, 240)
   Next H
   
   Set RecentColours = New cCollection
   
   Set RSlider = Widgets.Add(New cwSlider, "RSlider", 335, 10, 104, 18)
   Set GSlider = Widgets.Add(New cwSlider, "GSlider", 335, 42, 104, 18)
   Set BSlider = Widgets.Add(New cwSlider, "BSlider", 335, 74, 104, 18)

   Set ColorChooser = Widgets.Add(New cwColorChooser, "ColorChooser", 0, 30, 280, 280)
   Set SSlider = Widgets.Add(New cwSlider, "SatSlider", 20, 316, 240, 24)

   SSlider.Init "Sat", 0, 240, 240
   RSlider.Init "", 0, 255, 255
   GSlider.Init "", 0, 255, 0
   BSlider.Init "", 0, 255, 0

End Sub
Private Sub UserControl_Terminate()
   Set RecentColours = Nothing
   If Not Root Is Nothing Then
      Root.Widgets.RemoveAll
      Root.Disconnect
   End If
End Sub

