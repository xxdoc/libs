VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Piechart Example"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   510
      Left            =   9045
      TabIndex        =   15
      Top             =   3690
      Width           =   1230
   End
   Begin VB.PictureBox picLegend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   4095
      ScaleHeight     =   4215
      ScaleWidth      =   2160
      TabIndex        =   14
      Top             =   45
      Width           =   2160
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00CBA461&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      ScaleHeight     =   225
      ScaleWidth      =   1245
      TabIndex        =   12
      Top             =   2955
      Visible         =   0   'False
      Width           =   1245
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create"
      Height          =   375
      Left            =   8985
      TabIndex        =   10
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create"
      Height          =   330
      Left            =   6855
      TabIndex        =   9
      Top             =   3090
      Width           =   1065
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create A Pie Chart From A Random Number Array"
      Height          =   225
      Left            =   6660
      TabIndex        =   8
      Top             =   2835
      Width           =   4050
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8505
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "1"
      Top             =   1365
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   330
      Left            =   8985
      TabIndex        =   6
      Top             =   1335
      Width           =   960
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   6855
      TabIndex        =   5
      Top             =   1320
      Width           =   1545
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create A Pie Chart From An Array"
      Height          =   225
      Left            =   6660
      TabIndex        =   4
      Top             =   1035
      Width           =   3000
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create A Pie Chart From Intervals of 1"
      Height          =   225
      Left            =   6645
      TabIndex        =   3
      Top             =   225
      Value           =   -1  'True
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8010
      TabIndex        =   1
      Text            =   "0"
      Top             =   645
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Interval"
      Height          =   375
      Left            =   6885
      TabIndex        =   0
      Top             =   570
      Width           =   975
   End
   Begin PieChartOCX.Pie Pie1 
      Height          =   4110
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   7250
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   8070
      TabIndex        =   11
      Top             =   3105
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private i, e, f



Private Sub Command1_Click()
List1.AddItem Text1.Text
End Sub

Private Sub Command2_Click()
Dim chartpercents(0)
    f = Text2.Text + 1
    If f > 100 Then f = 1
    chartpercents(0) = f
    e = f
    Text2.Text = e
    e = 0
    i = 0
    Call Pie1.CreatePie(chartpercents, 300, 100, 20)
End Sub

Private Sub Command3_Click()
If Option3.Value = True Then
    Dim chartpercents(100)
    Randomize
    chartpercents(0) = 1 + Int(24 * Rnd)
    e = e + chartpercents(0)
    chartpercents(1) = 1 + Int(24 * Rnd)
    e = e + chartpercents(1)
    chartpercents(2) = 1 + Int(24 * Rnd)
    e = e + chartpercents(2)
    chartpercents(3) = 1 + Int(24 * Rnd)
    e = e + chartpercents(3)
    chartpercents(4) = 1 + Int(24 * Rnd)
    e = e + chartpercents(4)
    i = 4
    
    Do
        i = i + 1
        If e < 76 Then
            chartpercents(i) = 1 + Int(24 * Rnd)
            e = e + chartpercents(i)
        Else
            Exit Do
        End If
    Loop
    Label1.Caption = e
    e = 0
    i = 0
    Call Pie1.CreatePie(chartpercents, 1, 1, 1)
End If
End Sub

Private Sub Command4_Click()
If Option2.Value = True Then
    Dim chartpercents()
    ReDim chartpercents(List1.ListCount - 1)
    For x = 0 To List1.ListCount - 1
        chartpercents(x) = List1.List(x)
    Next
    Call Pie1.CreatePie(chartpercents, 300, 100, 40)
End If
End Sub

Private Sub Command5_Click()
    Form_Load
End Sub

Private Sub Form_Load()
'Pie1.NewColor RGB(97, 164, 203), RGB(61, 138, 184)
'Pie1.NewColor RGB(243, 243, 243), RGB(219, 219, 219)
'Pie1.Backcolor = RGB(255, 255, 255)
'Call Pie1.CreatePie(Array(10, 40, 20), 300, 100, 40)



'Call Pie2.CreatePie(Array(10, 40, 20, 30), 300, 100, 40)

Dim a(9)
Dim labels(9) As String
Pie1.ClearColors

For i = 0 To UBound(a)
    a(i) = 10
    labels(i) = "label " & i & "  (" & a(i) & ")"
Next

Pie1.CreatePie a, 200, 100, 20
Pie1.DrawLegend picLegend, labels



End Sub

Private Sub Pie1_MouseOutPiePiece(index As Variant, button As Variant, x As Variant, y As Variant)
Picture1.Visible = False
End Sub

Private Sub Pie1_MouseOverPiePiece(index As Variant, button As Variant, x As Variant, y As Variant)
    Picture1.Move Pie1.Left + (x * Screen.TwipsPerPixelX) + 300, Pie1.Top + (y * Screen.TwipsPerPixelY) + 300
    If Label2.Caption <> index Then Label2.Caption = index
    If Picture1.Visible = False Then Picture1.Visible = True
End Sub

