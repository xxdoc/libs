VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   540
      TabIndex        =   0
      Top             =   45
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ba As New CByteArray

Private Sub Form_Load()
    
    'test1
    push_pop_test
     
    'ba.pushHexStr "9090eb15"
    'ba.pushHexStr "90 90 eb 15"
    'MsgBox ba.hexdump
    
End Sub

Function push_pop_test()
     
    Dim b() As Byte
    
    ReDim b(3)
    b(0) = &H55
    b(1) = &H66
    b(2) = &H77
    b(3) = &H88
    
    ba.pushByte &H99
    ba.pushShort &H1122
    ba.pushLong &H11223344
    
    MsgBox Hex(ba.popLong)
    MsgBox Hex(ba.popShort)
    MsgBox Hex(ba.popByte)

    MsgBox ba.Length
    
    ba.Clear
    ba.pushBlock b()
    b() = ba.popBlock(4)
    ba.Clear
    ba.pushBlock b
    MsgBox ba.hexdump
    
End Function

Function test1()
    
     
    Dim b() As Byte
    
    ReDim b(3)
    b(0) = &H55
    b(1) = &H66
    b(2) = &H77
    b(3) = &H88
    
    
    ba.pushByte &H99
    ba.pushShort &H1122
    ba.pushLong &H11223344
    MsgBox ba.hexdump
    
    ba.pushBlock b()
    ba.pushStr "test"
    
    MsgBox ba.hexdump
    
    Dim i
    
    MsgBox Hex(ba.ReadShort(1))
    MsgBox Hex(ba.ReadLong(3))
 
    
End Function
