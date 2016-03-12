VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   270
      TabIndex        =   0
      Top             =   405
      Width           =   4020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    'most put together from existing code but not
    'everything has been tested in this context yet..
    Dim cs As New CString
    Dim pos As Long
    Dim b() As Byte
    Dim marker As String
    Dim pth As String
    
    pth = Environ("temp") & "\test.txt"
    
    b() = StrConv("testing testing!", vbFromUnicode, &H409)
    cs = "this is my text %3c:"
    
    With List1
        .AddItem cs.charAt(1)
        .AddItem cs.charCodeAt(1)
        .AddItem cs.indexOf("is")
        .AddItem cs.replace("text", "dog")
        .AddItem cs.substr(2, cs.indexOf(" "))
        .AddItem cs.endsWith("t")
        .AddItem cs.startsWith("thi")
        .AddItem cs.HexDump(, True)
        .AddItem cs.unescape()
        
        cs = "val='this is my val';val2='this is val2';"
        .AddItem cs.extract("'", "'", , pos)
        .AddItem cs.extract("'", "'", pos)
        
        cs.LoadFromHexString "6920736D656C6C2061206661727421"
        .AddItem cs.text
        
        cs.LoadFromBytes b()
        .AddItem cs
        
        Debug.Print cs.HexDump
        .AddItem cs.HexDump(b(), True)
        .AddItem cs.HexDump("test", True)
        
        Debug.Print cs.HexDump("1234567890", , 4, 4)
        
        Debug.Print cs.toHexString("test")
        Debug.Print cs.toHexString
        
        cs = "myTest('arg0');"
        pos = cs.findNextChar("'(|{}.)""", marker)
        .AddItem "first marker is " & marker & " at pos: " & pos
        
        cs = String(1000, "A")
        cs.SaveToFile pth
        cs.LoadFromFile pth
        .AddItem "Loaded " & cs.length & " bytes from file"
        
'        b() = cs.Compress()
'        marker = cs.DeCompress(b())
'        .AddItem "Compressed 1000 bytes to " & UBound(b) & " decompressed to: " & Len(marker)
        
        
        
    End With

End Sub
