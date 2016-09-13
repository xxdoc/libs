VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public file As String
Public sysHandle As Long
Public vbHandle As Long
Private buf() As Byte
Private pointer As Long

Public RedirectTo As RedirTypes
Public target As Variant
Private readPointer As Long

Public Enum RedirTypes
    rt_none = 0
    rt_memory = 1
    rt_ip = 2
    rt_Handle = 3
End Enum

Property Get isRedirected() As Boolean
    isRedirected = Not (RedirectTo = rt_none)
End Property

Property Get Data() As Byte()
    ReDim Preserve buf(pointer - 1)
    Data = buf
End Property

Sub AddData(lpBuf As Long, bufSize As Long)
    Dim size As Long
    
    If bufSize = 0 Then Exit Sub
    
    size = bufSize
    If size < &H1000 Then size = &H1000
    If size > &H1000 Then size = size + &H1000
    If AryIsEmpty(buf) Then ReDim buf(size)
    
    If pointer + bufSize > UBound(buf) Then
        size = UBound(buf) + bufSize + &H1000
        ReDim Preserve buf(size)
    End If
    
    CopyMemory ByVal VarPtr(buf(pointer)), ByVal lpBuf, bufSize
    pointer = pointer + bufSize
    
End Sub

Sub ResetData()
    readPointer = 0
    pointer = 0
    Erase buf
End Sub


Function ReadData(lpBuf As Long, length As Long) As Boolean
    
    If Not IsArray(target) Then Exit Function
    If AryIsEmpty(target) Then Exit Function
    
    Dim b() As Byte
    b() = target
    
    If readPointer + length > UBound(b) + 1 Then Exit Function
    
    CopyMemory ByVal lpBuf, ByVal VarPtr(b(readPointer)), length
    readPointer = readPointer + length
    ReadData = True
    
End Function
