VERSION 5.00
Begin VB.Form frmPie 
   AutoRedraw      =   -1  'True
   Caption         =   "Pie chart Demo"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPie 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   5175
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "&Draw"
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "frmPie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'create an object using the clas clsPie
Private PieGraph As New clsPie


Private Sub cmdDraw_Click()
    'Value is the value for the segment
    'Name is what the segment represents
    'Colour is the colour the segment is required in

                     'Value  'Name       'Colour
    PieGraph.AddSegment 1000, "Response", &HFF0000     'Blue
    PieGraph.AddSegment 1000, "Actual", &HFFFF&     'Yellow
    PieGraph.AddSegment 1000, "Delays", &HFF&       'Red
    PieGraph.AddSegment 100, "Test", &HFF00FF        'Violet
    
    'Draw the pie chart.
    PieGraph.DrawPie picPie.hdc, picPie.hwnd, True, "A Graph To Show The Breakdowns"
    
    'Clear the segments ready for redraw.
    PieGraph.Clear
End Sub

Private Sub Form_Load()
    cmdDraw_Click
End Sub
