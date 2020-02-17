VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCDef 
      Height          =   2865
      Left            =   11925
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   3075
      Width           =   4215
   End
   Begin VB.TextBox txtVBDef 
      Height          =   2865
      Left            =   11925
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":006F
      Top             =   75
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   9840
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   75
      Width           =   11715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'todo: parse vb/c structs into named CMemStruct's
'      add support for some of kaitai struct language? https://kaitai.io/

Private Type vbTest
    byte1 As Byte
    int1 As Integer
    long1 As Long
    cur1 As Currency
    blob1(14) As Byte
End Type

Private Sub Form_Load()
    
    Dim ms As New CMemStruct, errMsg As String
    Dim b() As Byte, tmp() As String
    
    'GoTo test4
    
    If Not ms.AddFields("byte1*b,int1*i,lng1*l,cur1*c,blob1*15") Then
        Text1 = ms.lastErr
        Exit Sub
    End If
    
    Dim f As Long
    Dim bb() As Byte
    Dim vbt As vbTest
    
    f = FreeFile
    Open App.Path & "\test.bin" For Binary As f
    ReDim bb(LOF(f) - 1)
    Get f, , bb()
    Get f, 1, vbt
    Close f
    
    addText "Raw file:\n" & HexDumpB(bb) & "\n\n"
    
    push tmp, "byte1 = " & Hex(vbt.byte1)
    push tmp, "int1 = " & Hex(vbt.int1)
    push tmp, "long1 = " & Hex(vbt.long1)
    push tmp, "cur1 = " & CurToHex(vbt.cur1)
    push tmp, "blob1 = " & HexDumpB(vbt.blob1)
    addText "vb UDT compatability test: \n" & Join(tmp, vbCrLf) & "\n"
    
    ms.LoadFromFile , App.Path & "\test.bin"
    b() = ms.toBytes()
        
    addText ms.dump(True)
    addText "\nFull struct .toBytes() hexDump:\n" & HexDumpB(b)
    
    ms.SaveToFile , App.Path & "\test2.bin"
    
    addText "\nOffsetOf lng1 = " & Hex(ms.offsetOf("lng1"))
    addText "OffsetOf blob1 = " & Hex(ms.offsetOf(5))
    addText "Structure Size = " & Hex(ms.size)
    
    If Not ms.field("blob1").SetValue("new blob2!") Then
        addText "Error setting new blob1 in test2.bin: " & ms.lastErr
    End If
    
    ms.SaveToFile 'we dont modify file offset or handle so it will dump to next address of cur file
    ms.LoadFromFile ms.size 'now we load the second structure from this file
    addText "second struct from text2.bin.blob.asString() = " & ms.field("blob1").asString()
    
    ms.LoadFromFile , App.Path & "\test.bin"
    ms.field("cur1") = 3.14 'using default property value =
    
    ms.SaveToFile , App.Path & "\test3.bin"
    ms.LoadFromFile , App.Path & "\test3.bin"
    
    addText "\nNew cur1 value reloaded from file = " & ms.field("cur1") 'using default prop get .value
    
    b(0) = &H80
    b(UBound(b)) = Asc("a")
    ms.fromBytes b
    addText "Loaded from modified byte buffer:\n " & ms.dump(True)
    
    addText "\nbyte1 as binary: " & Hex(ms.field(1).Value) & " = " & ms.field(1).toBinary()
    addText "byte1.clearBit(7).value = " & ms.field(1).clearBit(7).Value
    addText "byte1.setBit(7).setBit(6).toBinary() = " & ms.field(1).setBit(7).setBit(6).toBinary()
    
    Dim hs As String
    
    hs = ms.toHexString()
    addText "\nHexstring:" & hs
    
    If Not ms.fromHexString(Replace(hs, "11", "88")) Then
        addText "Error convertine from hex string! " & ms.lastErr
    Else
        addText "\nDumped from modified hex string: \n" & ms.dump(True)
    End If
        
    Exit Sub
    
test4:
    Set ms = New CMemStruct
    
'    Private Type PACKAGE_ID
'      reserved As Long
'      processorArchitecture  As Long
'      revision As Integer
'      build As Integer
'      minor As Integer
'      major As Integer
'      name  As Long
'      publisher  As Long
'      resourceId  As Long
'      publisherId  As Long
'    End Type

    Dim fields As String
    fields = "reserved*l,arch*l,rev*i,build*i,min*i,maj*i,name*l,pub*l,resource*l,pubID*l"
    
    If Not ms.AddFields(fields) Then
        Text1 = ms.lastErr
        Exit Sub
    End If
    
    ms.LoadFromFile , App.Path & "\test4.bin"
    Debug.Print ms.dump
    
    
    
    
End Sub


Sub addText(t)
    Text1 = Text1 & vbCrLf & Replace(t, "\n", vbCrLf)
    Text1.SelStart = Len(Text1)
End Sub
