VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUclCompression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This is an example of using the UCL Compression library in VB
'
'This VB file compression implementation is really only designed for
'use with small files because the entire block is compressed in memory
'all at once. Also a better implementation would be to embed some
'uncompressed size buffer in the compressed file because you cant really
'know how much it was compressed without it.
'
'Look at the sample uclpack from the ucl source distribution from the
'authors original C source.
'
'These limitations are because of a quick VB wrapper designed for a specific
'minimal use and are not limitations of the underlying UCL library.
'
'UCL is a GPL open source library copyright 2002 Markus Franz Xaver Johannes Oberhumer
'more information on the UCL compression library as well as its full source can
'be downloaded from the following site:
'
'http://www.oberhumer.com/opensource/ucl/
'
'This VB sample and C wrapper dll are also released GPL and the author makes
'no warrenties and can offer no support on its use.
'
'By using this library and implementation you take full responsible for any
'possibilities of damages, data loss, bugs, or crashs. In no event shall the author
'be liable for any damages.




Private Declare Sub SetLevel Lib "ucl_dll.dll" (ByVal x As Long)
Private Declare Sub uclLastError Lib "ucl_dll.dll" Alias "LastError" (buf As Byte)
Private Declare Function init Lib "ucl_dll.dll" Alias "Init" () As Long
Private Declare Function uclCompress Lib "ucl_dll.dll" Alias "Compress" (inBuf As Byte, outBuf As Byte, ByVal inlen As Long, outlen As Long) As Long
Private Declare Function uclDeCompress Lib "ucl_dll.dll" Alias "DeCompress" (inBuf As Byte, outBuf As Byte, ByVal in_len As Long, out_len As Long) As Long


Public LastError As String


Private Sub Class_Initialize()
     If init() <> 1 Then SetError
End Sub

Private Sub SetError(Optional msg As String)
        
        If Len(msg) > 0 Then
            LastError = msg
        Else
            Dim b(50) As Byte
            uclLastError b(0)
            LastError = StrConv(b(), vbUnicode)
        End If
        
End Sub

Property Let Level(x As Long)
    
    If x > 10 Or x < 1 Then
        SetError "Compression Level Must be 1-10"
        Exit Property
    End If
    
    SetLevel x
    
End Property

Function Compress(buf() As Byte, bout() As Byte) As Boolean
    
    On Error GoTo hell
    
    Dim compressedLen As Long
    Dim inlen As Long
    Dim f As Long
    
    ReDim bout(UBound(buf) * 2)

    inlen = UBound(buf) + 1
    compressedLen = UBound(bout)
     
    If uclCompress(buf(0), bout(0), inlen, compressedLen) < 0 Then GoTo hell
    
    ReDim Preserve bout(compressedLen) '- 1)
    
    Compress = True
    
    Exit Function
hell:

    If Err.Number > 0 Then
        SetError Err.Description
    Else
        SetError
    End If
    
    ReDim bout(0)
    
End Function

Function Decompress(buf() As Byte, padSz As Long, bout() As Byte) As Boolean
    
    On Error GoTo hell
    
    ReDim bout(padSz)
    
    If uclDeCompress(buf(0), bout(0), UBound(buf), padSz) < 0 Then GoTo hell
        
    ReDim Preserve bout(padSz - 1)
    Decompress = True

Exit Function
hell:

        If Err.Number > 0 Then
            SetError Err.Description
        Else
            SetError
        End If
       
        ReDim bout(0)
        
        
End Function

Function CompressFile(fpath As String, outFile As String, Optional Overwrite As Boolean = False) As Boolean
    
    If Not FileExists(fpath) Then Exit Function
    
    If FileExists(outFile) Then
        
        If Not Overwrite Then
            SetError "File Exists"
            Exit Function
        End If
        
        On Error Resume Next
        Kill outFile
        
        If Err.Number > 0 Then
            SetError "Could not Delete file:" & outFile
            Exit Function
        End If
        
    End If
    
    Dim f As Long
    Dim b() As Byte
    Dim bout() As Byte
    
    f = FreeFile
    Open fpath For Binary As f
    ReDim b(LOF(f) - 1)
    Get f, , b()
    Close f
    
    If Not Compress(b, bout) Then Exit Function
        
    Open outFile For Binary As f
    Put f, , bout()
    Close f
    
    CompressFile = True
    
End Function

Function DeCompressFile(fpath As String, outFile As String, Optional Overwrite As Boolean = False, Optional padSize As Long = 2000) As Boolean
    
    If Not FileExists(fpath) Then Exit Function
    
    If FileExists(outFile) Then
        
        If Not Overwrite Then
            SetError "File Exists"
            Exit Function
        End If
        
        On Error Resume Next
        Kill outFile
        
        If Err.Number > 0 Then
            SetError "Could not Delete File"
            Exit Function
        End If
            
    End If
    
    Dim f As Long
    Dim b() As Byte
    Dim bout() As Byte
    
    f = FreeFile
    Open fpath For Binary As f
    ReDim b(LOF(f) - 1)
    Get f, , b()
    Close f
    
    If Not Decompress(b, padSize, bout) Then Exit Function
    
    Open outFile For Binary As f
    Put f, , bout()
    Close f
    
    DeCompressFile = True
    
End Function



Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

