VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      Top             =   45
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
        
        cs = "test" & Chr(5) & Chr(5)
        .AddItem cs
        .AddItem cs.endsWith(Chr(5))
        cs.stripFromEnd Chr(5)
        .AddItem cs
        .AddItem cs.endsWith("est")
        
        If cs.LoadFromWeb("http://sandsprite.com/tools.php") Then
            .AddItem cs
        End If
        
        cs = "line0 \n line1 \n line2 \n line3"
        cs = cs.replace("\n", vbCrLf)
        .AddItem cs.getLine(0)
        .AddItem cs.getLine(2)
        .AddItem cs.getLine(7)
              
        cs = "test " & vbCr & vbLf & vbLf & " " & Chr(0)
        cs.stripAnyFromEnd vbLf, vbCr, " ", Chr(0)
        .AddItem cs & " (len: " & cs.length & ")"
        
        .AddItem cs.sprintf("number  %08x", &HCC)
        
    End With

End Sub
