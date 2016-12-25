VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

 Private Const Z_FINISH As Long = 4

 Public Enum ZLIB_CompressionLevelConstants
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1
    Z_BEST_COMPRESSION = 9
    Z_DEFAULT_COMPRESSION = (-1)
 End Enum

 Private Type zStream
    next_in As Long
    avail_in As Long
    total_in As Long
    next_out As Long
    avail_out As Long
    total_out As Long
    Msg As Long
    state As Long
    zalloc As Long
    zfree As Long
    opaque As Long
    data_type As Long
    adler As Long
    reserved As Long
 End Type

 Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
 Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
 Private Declare Function deflate Lib "zlib.dll" (vStream As zStream, Optional ByVal vflush As Long = Z_FINISH) As Long
 Private Declare Function deflateEnd Lib "zlib.dll" (vStream As zStream) As Long
 Private Declare Function deflateInit Lib "zlib.dll" Alias "deflateInit_" (strm As zStream, ByVal level As Long, ByVal version As String, ByVal stream_size As Long) As Long
 Private Declare Function inflate Lib "zlib.dll" (vStream As zStream, Optional ByVal vflush As Long = 1) As Long
 Private Declare Function inflateEnd Lib "zlib.dll" (vStream As zStream) As Long
 Private Declare Function inflateInit Lib "zlib.dll" Alias "inflateInit_" (strm As zStream, ByVal version As String, ByVal stream_size As Long) As Long
 Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
 Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

 Private msVersion As String
 Private mnChunkSize As Long
 Private hZLib As Long
 Private Const LANG_US As Long = &H409


Private Function ensureZLIB() As Boolean
    
    On Error Resume Next
    
    If hZLib <> 0 Then
        ensureZLIB = True
        Exit Function
    End If
    
    Dim pth As String, b() As Byte, f As Long
    
    pth = App.path & "\zlib.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\zlib.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\..\zlib.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\..\..\zlib.dll"
    
    If Not FileExists(pth) Then

        pth = App.path & "\zlib.dll"
        b() = LoadResData("ZLIB", "DLLS")
        If AryIsEmpty(b) Then
            MsgBox "Failed to find zlib.dll in resource?"
            Exit Function
        End If
        
        f = FreeFile
        Open pth For Binary As f
        Put f, , b()
        Close f
        
        'MsgBox "Dropped zlib.dll to: " & pth & " - Err: " & Err.Number
    End If
    
    If Not FileExists(pth) Then
        MsgBox "Failed to write zlib.dll to disk from resource?"
        Exit Function
    End If
        
    hZLib = LoadLibrary(pth)
        
    If hZLib = 0 Then
        MsgBox "Failed to load zlib.dll library?"
        Exit Function
    End If
    
    ensureZLIB = True
    
End Function

 
Public Property Get ZLIB_ChunkSize() As Long
    If mnChunkSize = 0 Then
        mnChunkSize = &H10000
    End If
    ZLIB_ChunkSize = mnChunkSize
End Property

Public Property Let ZLIB_ChunkSize(ByVal Value As Long)
    mnChunkSize = Value
End Property

Public Property Get ZLIB_Version() As String
    If LenB(msVersion) = 0 Then
        msVersion = "1.1.2.0"
    End If
    ZLIB_Version = msVersion
End Property

Public Property Let ZLIB_Version(ByRef Value As String)
    msVersion = Value
End Property

Public Function CompressData(ByRef vxbInput() As Byte, ByRef vxbOutput() As Byte, Optional vnStart As Long = 0, Optional vnMaxSize As Long = 0, Optional veCompressionLevel As ZLIB_CompressionLevelConstants = Z_DEFAULT_COMPRESSION) As Boolean
Dim tStream As zStream
Dim rc As Long
Dim xbCopy() As Byte
 
 Erase vxbOutput
 If Not ensureZLIB() Then Exit Function
 
 With tStream
 
    If deflateInit(tStream, veCompressionLevel, ZLIB_Version, Len(tStream)) = 0 Then
        CompressData = True
        
        CopyMemory rc, ByVal ArrPtr(vxbInput), 4
        If rc Then
            CopyMemory .avail_in, ByVal rc + 16, 4
            .avail_in = .avail_in - vnStart
        End If
        If .avail_in > 0 And vnStart < .avail_in Then
            If vnMaxSize <> 0 And vnMaxSize < .avail_in Then
                .avail_in = vnMaxSize
            End If
            .next_in = VarPtr(vxbInput(vnStart))

            CopyMemory rc, ByVal ArrPtr(vxbOutput), 4
            If rc Then
                CopyMemory rc, ByVal rc + 12, 4

                If rc + vnStart = .next_in Then
                    xbCopy = vxbInput
                    .next_in = VarPtr(xbCopy(vnStart))
                ElseIf vnStart Then
                    ReDim vxbOutput(vnStart - 1)
                    CopyMemory vxbOutput(0), vxbInput(0), vnStart - 1
                End If
            Else
                vxbOutput = vxbInput
            End If

            .avail_out = .avail_in + 12
 
            ReDim Preserve vxbOutput(.total_out - 1 + .avail_out + vnStart)
 
            .next_out = VarPtr(vxbOutput(vnStart + .total_out))

            CompressData = deflate(tStream, 4) = 1

            If .total_out Or vnStart Then
                ReDim Preserve vxbOutput(.total_out + vnStart - 1)
            Else
                Erase vxbOutput
            End If
        End If

        deflateEnd tStream
    End If
 End With
End Function

 Public Function UncompressData(ByRef vxbInput() As Byte, ByRef vxbOutput() As Byte, Optional vnStart As Long = 0, Optional vnMaxSize As Long = 0, Optional ByVal vnUncompressedSize As Long = 0) As Boolean
 Dim tStream As zStream
 Dim rc As Long
 Dim loop_counter As Long
 Dim warn_me As Boolean
 Dim mbr As VbMsgBoxResult
 Dim warn_size As Long
 Dim last_warn As Long
 Dim xbCopy() As Byte
 
 warn_me = True
 Erase vxbOutput
 If Not ensureZLIB() Then Exit Function

    With tStream
        
        If inflateInit(tStream, ZLIB_Version, Len(tStream)) = 0 Then
            UncompressData = True
            
            CopyMemory rc, ByVal ArrPtr(vxbInput), 4
            If rc Then
                CopyMemory .avail_in, ByVal rc + 16, 4
                .avail_in = .avail_in - vnStart
            End If
            If .avail_in > 0 And vnStart < .avail_in Then
                
                If vnMaxSize <> 0 And vnMaxSize < .avail_in Then
                    .avail_in = vnMaxSize
                End If
                .next_in = VarPtr(vxbInput(vnStart))
    
                CopyMemory rc, ByVal ArrPtr(vxbOutput), 4
                If rc Then
                    CopyMemory rc, ByVal rc + 12, 4
    
                If rc + vnStart = .next_in Then
                    xbCopy = vxbInput
                    .next_in = VarPtr(xbCopy(vnStart))
                ElseIf vnStart Then
                    ReDim xbDataOut(vnStart - 1)
                    CopyMemory vxbOutput(0), vxbInput(0), vnStart - 1
                End If
            ElseIf vnStart Then
                vxbOutput = vxbInput
            End If
    
            If vnUncompressedSize Then
                .avail_out = vnUncompressedSize
            Else
                .avail_out = .avail_in * 2
            End If

            Do
                ReDim Preserve vxbOutput(.total_out - 1 + .avail_out + vnStart)
               
                .next_out = VarPtr(vxbOutput(vnStart + .total_out))
    
                rc = inflate(tStream)
                If rc Then
                    UncompressData = rc > 0
                    Exit Do
                End If
    
             
                .avail_out = ZLIB_ChunkSize
                loop_counter = loop_counter + 1
                warn_size = Round((.total_out / 1024) / 1024, 0)
                DoEvents
                
                'decompression bomb detection..
                If (loop_counter Mod 500 = 0) Or _
                    (last_warn <> warn_size And warn_size Mod 10 = 0) _
                Then
                     last_warn = warn_size
                     'Form1.Caption = "Decompressing very large data     Current Size: " & warn_size & "mb    Hold ESC to abort this decompression.."
                     Sleep 10
                     DoEvents
                     
                     If GetAsyncKeyState(vbKeyEscape) <> 0 Then
                        'Form1.Caption = "Aborting the decompression of this stream!"
                        rc = 1
                     End If
                     
                    If warn_me Then
                        mbr = MsgBox("Possible decompression bomb detected." & vbCrLf & vbCrLf & _
                                "Current size is: " & warn_size & "mb" & vbCrLf & vbCrLf & _
                                "Choose ignore to disable this warning from showing again", vbAbortRetryIgnore)
                        If mbr = vbIgnore Then warn_me = False
                    End If
                    If mbr = vbAbort Then rc = 1

                End If
                
            Loop Until rc = 1
    
            If .total_out Or vnStart Then
                ReDim Preserve vxbOutput(.total_out + vnStart - 1)
            Else
                Erase vxbOutput
            End If
        End If
    
        inflateEnd tStream
    End If
 End With
End Function
 
 
Function DecompressString(ByVal s As String) As String
    Dim b() As Byte
    Dim bOut() As Byte
    
    b = StrConv(s, vbFromUnicode, LANG_US)
                
    UncompressData b(), bOut()
    
    If Not AryIsEmpty(bOut) Then
        DecompressString = StrConv(bOut, vbUnicode, LANG_US)
    Else
        DecompressString = s
        'MsgBox "Decompression Error"
    End If
                
End Function

Function CompressString(ByVal s As String) As String
    Dim b() As Byte
    Dim bOut() As Byte
    
    b = StrConv(s, vbFromUnicode, LANG_US)
                
    CompressData b(), bOut()
    
    If Not AryIsEmpty(bOut) Then
        CompressString = StrConv(bOut, vbUnicode, LANG_US)
    Else
        CompressString = Empty
        'MsgBox "Decompression Error"
    End If
                
End Function

Function CompressFile(ByVal pth As String, bOut() As Byte) As Boolean
    Dim b() As Byte
    Dim f As Long
    
    On Error Resume Next
    
    If Not FileExists(pth) Then Exit Function
    
    f = FreeFile
    Open pth For Binary As f
    ReDim b(LOF(f) - 1)
    Get f, , b()
    Close f
    
    If Err.Number <> 0 Then Exit Function
    
    CompressFile = CompressData(b, bOut)
    'If AryIsEmpty(bOut) Then Exit Function
    'CompressFile = True
       
End Function

Function DeCompressFile(ByVal pth As String, bOut() As Byte) As Boolean
    Dim b() As Byte
    Dim f As Long
    
    On Error Resume Next
    
    If Not FileExists(pth) Then Exit Function
    
    f = FreeFile
    Open pth For Binary As f
    ReDim b(LOF(f) - 1)
    Get f, , b()
    Close f
    
    If Err.Number <> 0 Then Exit Function
    
    DeCompressFile = UncompressData(b, bOut)
    'If AryIsEmpty(bOut) Then Exit Function
    'DeCompressFile = True
       
End Function

Function CompressFileToFile(ByVal pth As String, ByVal outFile As String) As Boolean
    
    Dim f As Long
    Dim bOut() As Byte
    On Error Resume Next
    
    If FileExists(outFile) Then Kill outFile
    If Not CompressFile(pth, bOut) Then Exit Function
    
    f = FreeFile
    Open outFile For Binary As f
    Put f, , bOut
    Close f
    CompressFileToFile = (Err.Number = 0)

End Function

Function DeCompressFileToFile(ByVal pth As String, ByVal outFile As String) As Boolean
    
    Dim f As Long
    Dim bOut() As Byte
    On Error Resume Next
    
    If FileExists(outFile) Then Kill outFile
    If Not DeCompressFile(pth, bOut) Then Exit Function
    
    f = FreeFile
    Open outFile For Binary As f
    Put f, , bOut
    Close f
    DeCompressFileToFile = (Err.Number = 0)

End Function



