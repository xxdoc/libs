VERSION 5.00
Object = "{42FF43C8-F73C-11D3-83DC-00A0CC355595}#1.0#0"; "BARCHART.OCX"
Begin VB.Form Form1 
   Caption         =   "BarChartX"
   ClientHeight    =   4530
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close Price Only"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Median Price"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin BarChartControl.BarChart BarChart1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1720
      BorderStyle     =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BarChart1_MouseMove(Price As Double, ObservationDate As String)
    If Price < 1 Then Exit Sub
    Caption = "Price: " & Price & " Date: " & ObservationDate
End Sub

Private Sub Check1_Click()
    DrawChart
End Sub

Private Sub DrawChart()
    Dim dbMain As Database
    Dim rsMain As Recordset
    Dim lngLoop As Long
    Dim MyArray As Variant

    Set dbMain = OpenDatabase(App.Path & "\Test.mdb")
    Set rsMain = dbMain.OpenRecordset("Soybeans")

    BarChart1.ClearChart

    ReDim MyArray(1 To rsMain.RecordCount, 1 To 5)

    Do While Not rsMain.EOF
        lngLoop = lngLoop + 1
        MyArray(lngLoop, 1) = rsMain!Open
        MyArray(lngLoop, 2) = rsMain!High
        MyArray(lngLoop, 3) = rsMain!low
        MyArray(lngLoop, 4) = rsMain!Close
        MyArray(lngLoop, 5) = rsMain!Day
        rsMain.MoveNext
        DoEvents
    Loop

    BarChart1.DrawBarChart MyArray, vbBlack

    'You can also draw a line chart
    If Check1 Then
        rsMain.MoveFirst
        lngLoop = 0
        ReDim MyArray(1 To rsMain.RecordCount)
        Do While Not rsMain.EOF
            lngLoop = lngLoop + 1
            MyArray(lngLoop) = (rsMain!High + rsMain!low) / 2
            rsMain.MoveNext
            DoEvents
        Loop
        BarChart1.DrawLineChart MyArray, vbBlue
    End If

    If Check2 Then
        BarChart1.ClearChart
        rsMain.MoveFirst
        lngLoop = 0
        ReDim MyArray(1 To rsMain.RecordCount)
        Do While Not rsMain.EOF
            lngLoop = lngLoop + 1
            MyArray(lngLoop) = rsMain!Close
            rsMain.MoveNext
            DoEvents
        Loop
        BarChart1.DrawLineChart MyArray, vbRed
    End If

End Sub

Private Sub Check2_Click()
    DrawChart
End Sub

Private Sub Form_Load()
    Call Check1_Click
End Sub

Private Sub Form_Resize()
    BarChart1.Width = Width - 120
    BarChart1.Height = Height - 400
End Sub
