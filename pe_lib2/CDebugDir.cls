VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDebugDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type debugDir
    Characteristics As Long
    timestamp As Long
    major As Integer
    minor As Integer
    dbgtype As Long
    sizeofData As Long
    adrRawData As Long
    ptrRawData As Long
End Type

Public Characteristics As Long
Public timestamp As Long
Public TimeStampStr As String
Public major As Integer
Public minor As Integer
Public dbgtype As Long
Public sizeofData As Long
Public adrRawData As Long
Public ptrRawData As Long
Public pdbPath As String

Private m_data() As Byte

Public Property Get Data() As Byte()
    Data = m_data
End Property


Friend Function loadSelf(pe As CPEEditor)
    
    Dim rvaDebug As Long
    Dim rawDebug As Long
    Dim f As Long
    Dim dd As debugDir
    Dim tmp As String
    Dim a As Long, b As Long
    
    On Error GoTo hell

    If Not pe.isLoaded Then Exit Function
    
    rvaDebug = pe.OptionalHeader.ddVirtualAddress(Debug_Data)
    If rvaDebug = 0 Then Exit Function
    
    rawDebug = pe.RvaToOffset(rvaDebug)
    If rawDebug = 0 Then Exit Function
    
    f = FreeFile
    Open pe.LoadedFile For Binary As f
    Get f, rawDebug + 1, dd
    
    Me.Characteristics = dd.Characteristics
    Me.adrRawData = dd.adrRawData
    Me.dbgtype = dd.dbgtype
    Me.major = dd.major
    Me.minor = dd.minor
    Me.ptrRawData = dd.ptrRawData
    Me.sizeofData = dd.sizeofData
    Me.TimeStampStr = timeStampToDate(dd.timestamp)
    Me.timestamp = dd.timestamp
    
    If dd.sizeofData > 0 And dd.sizeofData < &H500000 Then
        ReDim m_data(dd.sizeofData)
        Get f, dd.ptrRawData + 1, m_data()
    End If
    
    Close f
    f = 0
    
    tmp = StrConv(m_data, vbUnicode, &H409)
    a = InStr(1, tmp, ".pdb", vbTextCompare)
    If a < 1 Then Exit Function
    a = a + 4
    
    b = InStrRev(tmp, ":\", a)
    If b < 1 Then Exit Function
    b = b - 1
    
    pdbPath = Mid(tmp, b, a - b)
      
    Exit Function
hell:
    On Error Resume Next
    If f <> 0 Then Close f
    
End Function

