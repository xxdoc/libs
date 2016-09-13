VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back to UDT"
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdSetHooks 
      Caption         =   "Set Hooks"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test serialize UDT to memory"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   6480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3360
      Width           =   9135
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type test
    signature As Long
    buf As String
    signature2 As Long
    v As Variant
End Type

Dim seralized() As Byte

Private Sub cmdSetHooks_Click()
    If Not Init(List1) Then
        Command2.Enabled = False
        Command1.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    
    Dim f As Long
    Dim udtTest As test
    
    If Not mSer.isInit Then
        MsgBox "Initilize first"
        Exit Sub
    End If
    
    With udtTest
        .signature = &H11223344
        .buf = "test string"
        .signature2 = &H99887766
        .v = Array(1, "blah!", 3.14)
    End With

    f = mSer.StartSeralize()
    
    'MsgBox "putting structure!"
    Put f, , udtTest
    seralized() = mSer.EndSeralize()
    Text1 = HexDump(StrConv(seralized, vbUnicode, LANG_US))
    
End Sub

Private Sub Command2_Click()

    Dim f As Long
    Dim udtTest As test
    
    If Not mSer.isInit Then
        MsgBox "Initilize first"
        Exit Sub
    End If

    f = mSer.StartSeralize()
    mSer.SetReadBuffer seralized()
    
    'msgbox "reading structure!"
    Get f, , udtTest
    mSer.EndSeralize
    
    On Error Resume Next
    Dim tmp() As String
    
    push tmp, "Signature: " & Hex(udtTest.signature)
    push tmp, "Buf: " & udtTest.buf
    push tmp, "Signature2: " & Hex(udtTest.signature2)
    
    If IsArray(udtTest.v) Then
         push tmp, "v: " & Join(udtTest.v, ", ")
    Else
        push tmp, "error v is not array: " & TypeName(udtTest.v)
    End If
    
    Text1 = Join(tmp, vbCrLf)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mSer.isInit Then RemoveAllHooks
End Sub
