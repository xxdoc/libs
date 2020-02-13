VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13215
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   11955
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Sub Form_Load()

    Dim rh As New CRichHeader 'todo: clear_data no counts?
    Dim hash As New CWinHash 'dzrt reference (CRichHeader has no dependancies)
    
    'Call test_longs_to_bytes
    'Exit Sub
    
    If rh.Load("D:\_code\libs\pe_lib2\_sppe2.dll") Then
        'yara workbench:
        ' pe.dbg(hash.md5(pe.rich_signature.clear_data)) = 869487aa2b9d84eb86208f66ccbe7dd0
        Debug.Print hash.HashBytes(rh.clearData)       ' = 869487AA2B9D84EB86208F66CCBE7DD0
        Debug.Print hash.HashString(rh.strClearData) ' = 869487AA2B9D84EB86208F66CCBE7DD0
    End If
    
    Text1 = rh.dump 'ok if errors
   
   
End Sub


Function test_longs_to_bytes()
    
    Dim uPack() As Long, size As Long, b() As Byte
    
    ReDim uPack(1)
    uPack(0) = &H11223344
    uPack(1) = &H55667788
    
    size = (UBound(uPack) + 1) * 4
    ReDim b(size - 1)
 
    CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(uPack(0)), size
    Debug.Print HexDump(b) 'dzrt reference
    
    
End Function
