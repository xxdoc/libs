VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // clsStream.cls - binary stream class
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Private Const GRANULARITY   As Long = &H400  ' // Glanularity of memory allocation

Private mBuffer()       As Byte ' // Local buffer
Private mSizeOfBuffer   As Long ' // Size of buffer
Private mSizeOfData     As Long ' // Size of data
Private mCurIndex       As Long ' // Current pointer

' // Size of data
Public Property Get Size() As Long
    Size = mSizeOfData
End Property

' // Current pointer
Public Property Get BufferPointer() As Long
    BufferPointer = mCurIndex
End Property
Public Property Let BufferPointer( _
                    ByVal Value As Long)
    If Value < 0 Then
        Err.Raise 5
        Exit Property
    End If

    mCurIndex = Value
    
End Property

' // Compress stream
Public Sub CompressStream()
    Dim szWorkSpace     As Long
    Dim workSpace()     As Byte
    Dim output()        As Byte
    Dim outputSize      As Long
    Dim returnedSize    As Long
        
    If mSizeOfData = 0 Then
        Err.Raise 7
    End If
    
    If RtlGetCompressionWorkSpaceSize(COMPRESSION_FORMAT_LZNT1, szWorkSpace, 0) Then
        Err.Raise 7
    End If
    
    outputSize = mSizeOfData * 2
    ReDim workSpace(szWorkSpace - 1)
    ReDim output(outputSize - 1)
    
    If RtlCompressBuffer(COMPRESSION_FORMAT_LZNT1, _
                         mBuffer(0), mSizeOfData, _
                         output(0), outputSize, 4096, _
                         returnedSize, workSpace(0)) Then
        Err.Raise 7
    End If
    
    ReDim mBuffer(returnedSize - 1)
    
    CopyMemory mBuffer(0), output(0), returnedSize
    
    mSizeOfData = returnedSize
    
End Sub

' // Write data to stream
Public Sub WriteData( _
           ByVal pData As Long, _
           ByVal dwSize As Long)
           
    ' // Check data
    If IsBadReadPtr(ByVal pData, dwSize) Then
        Err.Raise 7
        Exit Sub
    End If
    
    ' // Check size of buffer
    If mCurIndex + dwSize >= mSizeOfBuffer Then
        
        If Not ReAllocMemory(mCurIndex + dwSize + 1) Then
            Err.Raise 7
            Exit Sub
        End If
        
    End If
    
    ' // Copy data to buffer
    CopyMemory mBuffer(mCurIndex), ByVal pData, dwSize
    
    mCurIndex = mCurIndex + dwSize
    
    If mCurIndex > mSizeOfData Then
        mSizeOfData = mCurIndex
    End If
    
End Sub

' // Write stream
Public Sub WriteStream( _
           ByVal oStream As clsStream)
    Dim oldStreamPointer    As Long
    Dim locBuffer()         As Byte
    
    On Error GoTo error_handler
    
    Err.Clear
    
    ' // Save old pointer
    oldStreamPointer = oStream.BufferPointer
        
    If oStream.Size > 0 Then

        oStream.BufferPointer = 0
        
        ' // Allocate memory
        ReDim locBuffer(oStream.Size - 1)
        
        oStream.ReadData VarPtr(locBuffer(0)), oStream.Size
        
        ' // Write to stream
        Me.WriteData VarPtr(locBuffer(0)), oStream.Size
        
    End If
    
error_handler:
    
    oStream.BufferPointer = oldStreamPointer
    
    ' // If an error occured throw exception
    If Err.Number Then Err.Raise Err.Number
    
End Sub

' // Read data from stream
Public Sub ReadData( _
           ByVal pData As Long, _
           ByVal dwSize As Long)
    Dim dwActualDataCount   As Long
    
    ' // Check memory
    If IsBadWritePtr(ByVal pData, dwSize) Then
        Err.Raise 7
        Exit Sub
    End If
    
    ' // Check size of buffer
    If mCurIndex + dwSize >= mSizeOfBuffer Then
        dwActualDataCount = mSizeOfBuffer - mCurIndex
    Else
        dwActualDataCount = dwSize
    End If
    
    If dwActualDataCount > 0 Then
        ' // Copy data
        CopyMemory ByVal pData, mBuffer(mCurIndex), dwActualDataCount
        pData = pData + dwActualDataCount
        
    End If
    
    If dwSize > dwActualDataCount Then
        ' // Fill with zero
        ZeroMemory ByVal pData, dwSize - dwActualDataCount
        
    End If
    
    mCurIndex = mCurIndex + dwSize
    
End Sub

' // Realloc memory
Private Function ReAllocMemory( _
                 ByVal dwSize As Long) As Boolean
    Dim dwNewBufferSize As Long
    
    On Error GoTo error_handler
    
    If dwSize < 0 Then
        Exit Function
    End If
    
    If dwSize Mod GRANULARITY Then
        dwNewBufferSize = (dwSize \ GRANULARITY) * GRANULARITY + GRANULARITY
    Else
        dwNewBufferSize = dwSize
    End If
    
    If dwNewBufferSize = 0 Then
        ' // Erase buffer
        Erase mBuffer()
    Else
        ' // Try to realloc memory
        ReDim Preserve mBuffer(dwNewBufferSize - 1)
    End If
    
    mSizeOfBuffer = dwNewBufferSize
    ReAllocMemory = True
    
error_handler:
    
End Function

