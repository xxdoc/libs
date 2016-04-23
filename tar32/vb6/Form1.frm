VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tar32.dll VB6 demo"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optMethod 
      Caption         =   "BZIP"
      Height          =   285
      Index           =   2
      Left            =   3375
      TabIndex        =   11
      Top             =   945
      Width           =   1275
   End
   Begin VB.OptionButton optMethod 
      Caption         =   "GZIP"
      Height          =   285
      Index           =   1
      Left            =   2430
      TabIndex        =   10
      Top             =   945
      Width           =   1275
   End
   Begin VB.OptionButton optMethod 
      Caption         =   "Tar"
      Height          =   285
      Index           =   0
      Left            =   1620
      TabIndex        =   9
      Top             =   945
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Create"
      Height          =   330
      Left            =   90
      TabIndex        =   8
      Top             =   900
      Width           =   1410
   End
   Begin VB.CheckBox chkIgnoreDir 
      Caption         =   "Ignore Dir struct"
      Height          =   285
      Left            =   8460
      TabIndex        =   7
      Top             =   495
      Width           =   1770
   End
   Begin VB.CheckBox chkHideDialog 
      Caption         =   "Hide Dialog"
      Height          =   240
      Left            =   6795
      TabIndex        =   6
      Top             =   540
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear List"
      Height          =   330
      Left            =   8370
      TabIndex        =   5
      Top             =   990
      Width           =   1725
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "List"
      Height          =   330
      Left            =   90
      TabIndex        =   4
      Top             =   495
      Width           =   1410
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1575
      TabIndex        =   3
      Top             =   135
      Width           =   8610
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "CMD"
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      Height          =   2265
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   4410
      Width           =   10140
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   90
      TabIndex        =   0
      Top             =   1395
      Width           =   10140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'vb6 api declares for tar32.dll by Yoshioka Tsuneo(tsuneo@rr.iij4u.or.jp)
'
'you can pack tar/tar.gz/gz/bz2 and unpack tar/tar.gz/tar.Z/tar.bz2/gz/Z/bz2.
'unpacking tar/tar.gz/tar.Z is auto-detect.
'
'homepage:
'   http://openlab.ring.gr.jp/tsuneo/tar32/index-e.html


Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function TarGetVersion Lib "tar32" () As Integer
Private Declare Function Tar Lib "tar32" (ByVal hwnd As Long, ByVal cmd As String, ByVal buf As String, ByVal bufLen As Long) As Long
Private Declare Function TarCheckArchive Lib "tar32" (ByVal szArcFile As String, ByVal mode As Long) As Long
Private Declare Function TarOpenArchive Lib "tar32" (ByVal hwnd As Long, ByVal fname As String, ByVal mode As Long) As Long
Private Declare Function TarCloseArchive Lib "tar32" (ByVal handle As Long) As Long
Private Declare Function TarFindFirst Lib "tar32" (ByVal hwnd As Long, ByVal wildCard_NOT_IMPLEMENTED As String, ByRef ii As INDIVIDUALINFO) As Long
Private Declare Function TarFindNext Lib "tar32" (ByVal hwnd As Long, ByRef ii As INDIVIDUALINFO) As Long
Private Declare Function TarGetFileCount Lib "tar32" (ByVal szArcFile As String) As Long

Private Type INDIVIDUALINFO
    dwOriginalSize As Long
    dwCompressedSize As Long
    dwCRC As Long
    uFlag As Long
    uOSType As Long
    wRatio As Integer
    wDate As Integer
    wTime As Integer
    szFileName As String * 513
    dummy1 As String * 3
    szAttribute As String * 8
    szMode As String * 8
    'safety As String * 10
End Type

Dim f As String

Private Sub cmdClear_Click()
    List1.Clear
End Sub

Private Sub cmdCompress_Click()
    Dim buf As String, cmd As String
    
    If optMethod(0).Value Then 'tar only
        cmd = "-cvf c:\foo.tar """ & App.path & "\..\src\*.cpp"""
    ElseIf optMethod(1).Value Then 'gzip
        cmd = "-cvfz c:\foo.tgz """ & App.path & "\..\src\*.cpp"""
    Else 'bzip2
        cmd = "-cvfB c:\foo.tar.bz2 """ & App.path & "\..\src\*.cpp"""
    End If
        
    cmd = cmd & " --use-directory=0 --display-dialog=0"
    
    buf = String(1000, Chr(0))
    x = Tar(Me.hwnd, cmd, buf, LenB(buf))
    List1.AddItem x
    Text1 = Replace(buf, vbLf, vbCrLf)
    
End Sub

Private Sub cmdExtract_Click()
    Dim buf As String, cmd As String
    cmd = Text2
    If chkHideDialog.Value = 1 Then cmd = cmd & " --display-dialog=0"
    If chkIgnoreDir.Value = 1 Then cmd = cmd & " --use-directory=0"
    
    buf = String(1000, Chr(0))
    x = Tar(Me.hwnd, cmd, buf, LenB(buf))
    List1.AddItem x
    Text1 = Replace(buf, vbLf, vbCrLf)
End Sub

Private Sub cmdFind_Click()
    Dim ii As INDIVIDUALINFO
    Dim h As Long, ret As Long, cnt As Long
    
    h = TarOpenArchive(0, f, 0)
    If h = 0 Then
        List1.AddItem "Failed to open tar: " & f
        Exit Sub
    End If
    
    ret = TarFindFirst(h, Empty, ii) 'doesnt seem to honor wildcard?
    While ret <> -1
        cnt = cnt + 1
        List1.AddItem "Found file: " & ii.szFileName
        ret = TarFindNext(h, ii)
    Wend
    
    TarCloseArchive h
    List1.AddItem "Found " & cnt & " items"
    
End Sub

Private Sub Form_Load()

    Dim h As Long
    
    h = LoadLibrary("tar32.dll")
    If h = 0 Then h = LoadLibrary(App.path & "\tar32.dll")
    If h = 0 Then h = LoadLibrary(App.path & "\..\tar32.dll")
    
    If h = 0 Then
        MsgBox "tar32.dll not found?"
        Exit Sub
    End If
    
    f = GetParentFolder(App.path) & "\test.tar.gz"
    Text2 = "-xvf """ & f & """ -o C:\test\"
    
    List1.AddItem "Tar32.dll Version: " & TarGetVersion()
    List1.AddItem "Check Archive: " & TarCheckArchive(f, 0)
    List1.AddItem "File count: " & TarGetFileCount(f)

End Sub



Function GetParentFolder(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function
