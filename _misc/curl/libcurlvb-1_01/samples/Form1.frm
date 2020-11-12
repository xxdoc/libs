VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13365
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   13365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   12375
      TabIndex        =   10
      Top             =   630
      Width           =   645
   End
   Begin VB.CheckBox chkVerbose 
      Caption         =   "Verbose"
      Height          =   330
      Left            =   5580
      TabIndex        =   9
      Top             =   990
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.TextBox txtOutput 
      Height          =   3300
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4995
      Width           =   12300
   End
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   765
      TabIndex        =   5
      Top             =   1890
      Width           =   12345
   End
   Begin VB.CommandButton cmdDl 
      Caption         =   "Download"
      Height          =   465
      Left            =   11430
      TabIndex        =   4
      Top             =   1035
      Width           =   1770
   End
   Begin VB.TextBox txtSaveAs 
      Height          =   360
      Left            =   1215
      TabIndex        =   3
      Top             =   540
      Width           =   11040
   End
   Begin VB.TextBox txtUrl 
      Height          =   360
      Left            =   1215
      TabIndex        =   2
      Text            =   "http://sandsprite.com/tools.php"
      Top             =   135
      Width           =   10995
   End
   Begin VB.Label Label5 
      Caption         =   "Empty SaveAs path = mem only"
      Height          =   285
      Left            =   1260
      TabIndex        =   11
      Top             =   990
      Width           =   4155
   End
   Begin VB.Label Label4 
      Caption         =   "Output ( size if file download else mem download)"
      Height          =   375
      Left            =   45
      TabIndex        =   7
      Top             =   4455
      Width           =   6765
   End
   Begin VB.Label Label3 
      Caption         =   "Debug"
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   1620
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "Save As"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'uses libcurl to download files directly to the toFile location no cache used.

'  file:   vblibcurl.dll (std C dll)
'  file    vblibcurl.tlb (api declares and enums for above) - see modDeclares for all enums and partial declares
'  author: Jeffrey Phillips
'  date:   2.28.2005
 
'mods to vb test code  dzzie@yahoo.com

'I added a ref to dzrt for a couple funcs (lazy test code)
'   https://github.com/dzzie/libs/tree/master/dzrt

'(REMOVED tlb references)
'    this code also requires refs to a couple tlb files in source directory from original author. add reference browse select
'    because we use a typeLib we cant call download method in the module until the dll is loaded into memory?

Function initLib() As Boolean
    
    If hLib <> 0 And hLib2 <> 0 Then
        initLib = True
        Exit Function
    End If
    
    Dim base() As String, b
    Const dll = "libcurl.dll"
    
    push base, App.Path
    push base, App.Path & "\bin"
    push base, fso.GetParentFolder(App.Path)
    push base, fso.GetParentFolder(App.Path) & "\bin"
    
    For Each b In base
        hLib = LoadLibrary(b & "\" & dll)
        If hLib <> 0 Then
            List1.AddItem "Loaded " & b & "\" & dll
            hLib2 = LoadLibrary(b & "\" & "vb" & dll)
            If hLib2 = 0 Then
                List1.AddItem "Failed to load vbLibcurl.dll from same directory?!"
            Else
                List1.AddItem "Loaded vbLibCurl.dll from same directory."
            End If
            Exit For
        End If
    Next
        
    If hLib <> 0 And hLib2 <> 0 Then
        initLib = True
    Else
        List1.AddItem "Could not load libcurl.dll"
    End If

End Function

Private Sub cmdBrowse_Click()
    txtSaveAs = fso.dlg.OpenDialog(fso.GetSpecialFolder(sf_DESKTOP))
End Sub

Private Sub cmdDl_Click()
    
    On Error GoTo hell
    
    Dim x, mem As CMemBuffer
    
    List1.Clear
    txtOutput = Empty
    
    If Len(txtSaveAs) = 0 Then
        Set mem = Download(txtUrl, , True)
        If mem.size = 0 Then
            txtOutput = "No data headers were: " & vbCrLf & vbCrLf & Join(c2a(headers), "")
        Else
            txtOutput = mem.asString
        End If
    Else
        txtOutput = "File size was: " & Download(txtUrl, txtSaveAs, True)
        txtOutput = txtOutput & vbCrLf & "MD5: " & hash.HashFile(txtSaveAs)
    End If
    
    For Each x In modDzTest.debugMsg
        List1.AddItem x
    Next
    
    Exit Sub
hell:
    List1.AddItem "Error: " & Err.Description
End Sub

Private Sub Form_Load()

    If Not initLib() Then cmdDl.Enabled = False
    
    Dim b As New CMemBuffer, b1() As Byte, b2() As Byte, b3() As Byte
    
    b1 = StrConv("0123456789ABCDEF", vbFromUnicode, &H409)
    b2 = StrConv(String(16, "B"), vbFromUnicode, &H409)
    b3 = StrConv(String(16, "C"), vbFromUnicode, &H409)
    
    'b.appendBuf b1
    'b.appendBuf b2
    'b.appendBuf b3
    'txtOutput = "CMemBufTest: " & vbCrLf & HexDump(b.binData())
    
    b.memAppendBuf VarPtr(b1(0)), UBound(b1) + 1
    b.memAppendBuf VarPtr(b2(0)), UBound(b2) + 1
    b.memAppendBuf VarPtr(b3(0)), UBound(b3) + 1
    txtOutput = "CMemBufTest: " & vbCrLf & HexDump(b.binData())

    
End Sub
