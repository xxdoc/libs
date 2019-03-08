VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   180
      Width           =   13395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'win8+

'LONG GetPackageFullName(
'  HANDLE hProcess,
'  UINT32 *packageFullNameLength,
'  PWSTR packageFullName
');

Private Declare Function GetPackageFullName Lib "kernel32" (ByVal hProcess As Long, ByRef packageFullNameLength As Long, ByRef packageFullName As Byte) As Long


'------------------------------------------------------------------
'LONG GetPackageId(
'  HANDLE hProcess,
'  UINT32 *bufferLength,
'  BYTE   *buffer
');

'buffer contains PACKAGE_ID header, PWSTR pointers point to unicode strings held in remainder of buffer following header
Private Declare Function GetPackageId Lib "kernel32" (ByVal hProcess As Long, ByRef bufferLength As Long, ByRef buffer As Byte) As Long

'typedef struct PACKAGE_ID { '16 bytes
'  UINT32          reserved;
'  UINT32          processorArchitecture;
'  PACKAGE_VERSION version;
'  PWSTR           name;          'mem pointer to within returned buffer
'  PWSTR           publisher;
'  PWSTR           resourceId;
'  PWSTR           publisherId;
'} PACKAGE_ID;

'typedef struct PACKAGE_VERSION { '8 bytes
'  union {
'    UINT64 Version;
'    struct {
'      USHORT Revision;
'      USHORT Build;
'      USHORT Minor;
'      USHORT Major;
'    } DUMMYSTRUCTNAME;
'  } DUMMYUNIONNAME;
'} PACKAGE_VERSION;

Private Type PACKAGE_ID
  reserved As Long
  processorArchitecture  As Long
  revision As Integer
  build As Integer
  minor As Integer
  major As Integer
  name  As Long
  publisher  As Long
  resourceId  As Long
  publisherId  As Long
End Type
'------------------------------------------------------------------




Private Sub Form_Load()

    Dim hProcess As Long
    Dim pid As Long
    Dim length As Long, rc As Long, buf() As Byte
    
    Const PROCESS_QUERY_LIMITED_INFORMATION = &H1000
    Const ERROR_INSUFFICIENT_BUFFER = 122 '(0x7A)
    Const APPMODEL_ERROR_NO_PACKAGE = 15700
    
    pid = CLng(InputBox("Enter pid of a store app to query: ", , 3900))
    
    hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    If hProcess = 0 Then
        d "error opening pid"
        Exit Sub
    End If
    
    rc = GetPackageFullName(hProcess, length, 0)
    
    If rc <> ERROR_INSUFFICIENT_BUFFER Then
        If rc = APPMODEL_ERROR_NO_PACKAGE Then
            d "this pid is not a windows store app"
        Else
            d "error GetPackageFullName returned " & rc
        End If
        Exit Sub
    End If
        
    
    d rc & ":" & length
    
    ReDim buf((length + 2) * 2)
    rc = GetPackageFullName(hProcess, length, buf(0))
    d rc
    
    'd HexDump(buf)
    d GetUniString(buf, 0)
    
    
    Dim pack As PACKAGE_ID
    Dim base As Long
    
    
    rc = GetPackageId(hProcess, length, 0)
    ReDim buf(length)
    base = VarPtr(buf(0))
    d Join(Array(rc, length), ":")
    
    rc = GetPackageId(hProcess, length, buf(0))
    d Join(Array(rc, length, Hex(base)), ":")
    'd HexDump(buf)
    
    CopyMemory ByVal VarPtr(pack), ByVal base, LenB(pack)
    
    d "name: " & Hex(pack.name - base) & " " & GetUniString(buf, pack.name - base)
    d "pub: " & Hex(pack.publisher - base) & " " & GetUniString(buf, pack.publisher - base)
    
    If pack.resourceId <> 0 Then
        d "resID: " & Hex(pack.resourceId - base) & " " & GetUniString(buf, pack.resourceId - base)
    Else
        d "resID: 0"
    End If
    
    d "pubid: " & Hex(pack.publisherId - base) & " " & GetUniString(buf, pack.publisherId - base)

   
End Sub

Function GetUniString(buf() As Byte, ByVal offset As Long) As String
    
    Dim tmp() As Byte
    Dim sz As Long
    Dim i As Long
    
    If offset < 0 Or offset > UBound(buf) Then Exit Function
    
    ReDim tmp(UBound(buf))
    
    For i = offset To UBound(buf)
        If buf(i) = 0 And buf(i + 1) = 0 Then Exit For
        If buf(i) <> 0 Then
            tmp(sz) = buf(i)
            sz = sz + 1
        End If
    Next
    
    If sz > 0 Then
        ReDim Preserve tmp(sz - 1)
        GetUniString = StrConv(tmp, vbUnicode)
    End If
        
End Function

Function d(x)
    Text1 = Text1 & x & vbCrLf
End Function
