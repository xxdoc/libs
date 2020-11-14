VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "vbLibCurl Demo"
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
   Begin VB.TextBox txtConnectTimeout 
      Height          =   360
      Left            =   7650
      TabIndex        =   16
      Text            =   "15"
      Top             =   1395
      Width           =   465
   End
   Begin VB.TextBox txtTimeout 
      Height          =   360
      Left            =   4365
      TabIndex        =   14
      Text            =   "0"
      Top             =   1395
      Width           =   510
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   420
      Left            =   10395
      TabIndex        =   12
      Top             =   1035
      Width           =   1140
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   5985
      TabIndex        =   11
      Top             =   1080
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   12375
      TabIndex        =   9
      Top             =   630
      Width           =   645
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
      Left            =   11700
      TabIndex        =   4
      Top             =   1035
      Width           =   1500
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
   Begin VB.Label Label7 
      Caption         =   "Connect timeout    (s)"
      Height          =   285
      Left            =   5535
      TabIndex        =   15
      Top             =   1440
      Width           =   3120
   End
   Begin VB.Label Label6 
      Caption         =   "Total DL Timeout    (s)"
      Height          =   285
      Left            =   2115
      TabIndex        =   13
      Top             =   1440
      Width           =   3390
   End
   Begin VB.Label Label5 
      Caption         =   "Empty SaveAs path = mem only"
      Height          =   285
      Left            =   1260
      TabIndex        =   10
      Top             =   990
      Width           =   4155
   End
   Begin VB.Label Label4 
      Caption         =   "Output"
      Height          =   375
      Left            =   45
      TabIndex        =   7
      Top             =   4455
      Width           =   915
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
Implements ICurlProgress

Private abort As Boolean
Dim fso As Object 'if they have dzrt.CFileSystem3 then browse button is available..

'uses libcurl to download files directly to the toFile location no cache used.

'  file:   vblibcurl.dll (std C dll)
'  file    vblibcurl.tlb (api declares and enums for above) - see modDeclares for all enums and partial declares
'  author: Jeffrey Phillips
'  date:   2.28.2005
 
'dzzie: 11.13.20
'   initLib() to find/load C dll dependencies on the fly from different paths
'   removed tlb references w/modDeclares.bas (all enums covered but not all api declares written yet)
'   added higher level framework around low level api
'   file progress, response object, abort, download to memory only
'   still more that could be done/cleaned up but this is all i need for now and keystrokes are limited :(


Private Sub cmdAbort_Click()
    abort = True
End Sub

Private Sub ICurlProgress_Header(obj As CCurlResponse, msg As String)
     List1.AddItem "header: " & msg
End Sub

Private Sub ICurlProgress_InfoMsg(obj As CCurlResponse, info As curl_infotype, msg As String)
    List1.AddItem "info: " & info & " (" & info2Text(info) & ") " & msg
End Sub

Private Sub ICurlProgress_Init(obj As CCurlResponse)
    On Error Resume Next
    If obj.DownloadLength > 0 Then pb.Max = obj.DownloadLength
End Sub

Private Sub ICurlProgress_Progress(obj As CCurlResponse)
    On Error Resume Next
    pb.Value = obj.BytesReceived
    If abort Then
        List1.AddItem "Aborting at user request..."
        obj.abort = True
    End If
End Sub

Private Sub ICurlProgress_Complete(obj As CCurlResponse)
    pb.Value = 0
    List1.AddItem "Download complete resp code: " & obj.ResponseCode & " time: " & obj.TotalTime
End Sub


Private Sub cmdBrowse_Click()
    On Error Resume Next 'late bound if found
    Const sf_DESKTOP As Long = &H0
    txtSaveAs = fso.dlg.OpenDialog(fso.GetSpecialFolder(sf_DESKTOP))
End Sub

Private Sub cmdDl_Click()

    On Error Resume Next
    
    Dim X, resp As CCurlResponse
    Dim totalTimeout As Long, connectTimeout As Long
        
    List1.Clear
    txtOutput = Empty
    abort = False
    
    totalTimeout = CLng(txtTimeout)
    If Err.Number <> 0 Then
        List1.AddItem "Invalid total timeout"
        Exit Sub
    End If
    
    connectTimeout = CLng(txtConnectTimeout)
    If Err.Number <> 0 Then
        List1.AddItem "Invalid connect timeout"
        Exit Sub
    End If
    
    On Error GoTo hell
    
    If Len(txtSaveAs) = 0 Then
        Set resp = Download(txtUrl, , Me, connectTimeout, totalTimeout)
        txtOutput = resp.dump & vbCrLf & vbCrLf & resp.memFile.asString
    Else
        Set resp = Download(txtUrl, txtSaveAs, Me, connectTimeout, totalTimeout)
        txtOutput = resp.dump
        'txtOutput = txtOutput & vbCrLf & "MD5: " & hash.HashFile(txtSaveAs)
    End If
    
    Exit Sub
hell:
    List1.AddItem "Error: " & Err.Description
End Sub

Private Sub Form_Load()

    If Not initLib() Then
        cmdDl.Enabled = False
        List1.AddItem "This demo requires vblibcurl.dll and libcurl.dll"
        List1.AddItem "https://sourceforge.net/projects/libcurl-vb/files/libcurl-vb/libcurl.vb%201.01/"
    End If
    
    'membuf test
    '--------------------------------
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
    '--------------------------------
 
    On Error Resume Next
    Set fso = CreateObject("dzrt.CFileSystem3")
    If Err.Number <> 0 Then
        cmdBrowse.Enabled = False
        List1.AddItem "dzrt.dll not found browse file disabled"
    End If
    
    
End Sub



