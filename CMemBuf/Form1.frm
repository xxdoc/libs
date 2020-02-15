VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim mb As New CMemBuf
    Dim i As Integer
    Dim l As Long
    Dim ii(4) As Integer
    Dim ss() As String
    Dim strStart As Long
    
'output:
    '000000   22 11 33 44 55 66 22 11 44 33 66 55 88 77 AA       .".3DUf".D3fU.w.
    '000010   99 74 65 73 74 69 6E 67 21 00 04 00 74 68 69 73    .testing!...this
    '000020   02 00 69 73 02 00 6D 79 06 00 73 74 72 69 6E 67    ..is..my..string
    '000030   34 34                                              44
    '55443311
    'testing!
    'this
    'is
    'my
    'string
    '44

    i = &H1122
    l = &H66554433
    
    ii(0) = &H1122
    ii(1) = &H3344
    ii(2) = &H5566
    ii(3) = &H7788
    ii(4) = &H99AA
    
    ss = Split("this is my string", " ")
    
    mb.write_ i
    mb.write_ l
    mb.write_ ii
    
    strStart = mb.curOffset
    mb.writeStr "testing!"
    mb.writeStr ss, , mb_bstr
    mb.writeStr 44, , mb_fixedSize
    
    Debug.Print mb.HexDump
    
    mb.read l, 2
    Debug.Print Hex(l)
    
    Debug.Print mb.readStr(strStart)
    Debug.Print mb.readStr(, mb_bstr)
    Debug.Print mb.readStr(, mb_bstr)
    Debug.Print mb.readStr(, mb_bstr)
    Debug.Print mb.readStr(, mb_bstr)
    Debug.Print mb.readStr(, mb_fixedSize, 2)
    
    
    
    
End Sub
