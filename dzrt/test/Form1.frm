VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   1740
      TabIndex        =   1
      Top             =   1350
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   3000
      TabIndex        =   0
      Top             =   630
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Dim s As New StringEx
    Const x As Boolean = False
    
    d = esc(" \r\n\t This is my string  \t\r\n ")
    
    s = d
    
    Debug.Print unesc("'" & s & "'")
    Debug.Print unesc("'" & s.RTrim(x) & "'")
    Debug.Print unesc("'" & s.LTrim(x) & "'")
    
End Sub

Private Sub Command2_Click()

    Dim x()
    Dim ok As Boolean

    tmp = String(4000, "A")
    push x, "InitHash: " & hash.HashString(CStr(tmp)) & " Len: " & Len(tmp)
    c = Compress(tmp, True)
    push x, "Compressed size: " & Len(c)
    
    If Not DeCompress(c, d, True) Then
        MsgBox "Decompress failed"
        Exit Sub
    End If
    
    push x, "Decompressed hash: " & hash.HashString(CStr(d)) & " Len: " & Len(d)
    
    MsgBox Join(x, vbCrLf)
End Sub

Private Sub Form_Load()
    

    Dim fp  As CFileProperties
    'c:\windows\notepad.exe
    Set fp = fso.FileProperties("D:\_back_me_up\OSFILES\WIN10\0007A2C457A3823F930BA1FFE14FE18A", , "CompanyName,fileversion")
    MsgBox fp.CompanyName
    MsgBox fp.CustomFields("fileversion")
    

    Exit Sub

    Dim s As New StringEx

    s = "c:\windows\system32"
    s.CollapseConstants

    MsgBox s
    

    
        

    
End Sub

Function unesc(x)
    unesc = Replace(x, vbTab, "\t")
    unesc = Replace(unesc, vbCr, "\r")
    unesc = Replace(unesc, vbLf, "\n")
End Function

Function esc(x)
    esc = Replace(x, "\t", vbTab)
    esc = Replace(esc, "\r", vbCr)
    esc = Replace(esc, "\n", vbLf)
End Function
