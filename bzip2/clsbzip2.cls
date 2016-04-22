VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBzip2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'4.20.2016 - simplified by david zimmer <dzzie@yahoo.com>

'11.24.2003 altered by mario sepulveda and posted at vbforums.com
'BZIP2 homepage: http://sources.redhat.com/bzip2/

'comments below this line are from the original version found on the web
' ***
' * libbz2.dll calling interface for VB
' *   coded by Arnout de Vries, Relevant Soft- & Mindware
' *   24 jan 2001
' *   22 apr 2002, small update suggested by Alexander Feuster
' *     When using Win98SE you need the normal dll API
' *     When using Win2K you need the Alt variant
' *     Used simple error trapping to correct.
' *
' *   Enjoy and use it as much as possible
' *
' * BZIP2 homepage: http://sourceware.cygnus.com/bzip2/
' * from the webpage:
' *    What is bzip2?
' *    bzip2 is a freely available, patent free (see below), high-quality data compressor.
' *    It typically compresses files to within 10% to 15% of the best available techniques
' *    (the PPM family of statistical compressors), whilst being around twice as fast at
' *    compression and six times faster at decompression.
' ***

'Constants
Private Enum bzErrorCodes
    BZ_OK = 0
    BZ_RUN_OK = 1
    BZ_FLUSH_OK = 2
    BZ_FINISH_OK = 3
    BZ_STREAM_END = 4
    BZ_SEQUENCE_ERROR = (-1)
    BZ_PARAM_ERROR = (-2)
    BZ_MEM_ERROR = (-3)
    BZ_DATA_ERROR = (-4)
    BZ_DATA_ERROR_MAGIC = (-5)
    BZ_IO_ERROR = (-6)
    BZ_UNEXPECTED_EOF = (-7)
    BZ_OUTBUFF_FULL = (-8)
    BZ_CONFIG_ERROR = (-9)
End Enum

Private Const BZh As String * 3 = "BZh"

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Private Declare Function z2CompressAlt Lib "libbz2.dll" Alias "BZ2_bzBuffToBuffCompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal blockSize100k As Long, ByVal Verbosity As Long, ByVal workFactor As Long) As Long
Private Declare Function z2DecompressAlt Lib "libbz2.dll" Alias "BZ2_bzBuffToBuffDecompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal Small As Long, ByVal Verbosity As Long) As Long

'used for win98 support only
'Private Declare Function z2Compress Lib "libbz2.dll" Alias "bzBuffToBuffCompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal blockSize100k As Long, ByVal Verbosity As Long, ByVal workFactor As Long) As Long
'Private Declare Function z2Decompress Lib "libbz2.dll" Alias "bzBuffToBuffDecompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal Small As Long, ByVal Verbosity As Long) As Long


Public Enum eCompressionLevels
    cl1Fast = 1
    cl2 = 2
    cl3 = 3
    cl4 = 4
    cl5 = 5
    cl6 = 6
    cl7 = 7
    cl8 = 8
    clBest = 9
End Enum

Private m_lDecompressedSize As Long '-In Bytes
Private m_lCompressedSize As Long   '-In Bytes

Public LastError As String
 
Private Function BZErrorCodeToString(ec As bzErrorCodes)

    Select Case ec
        Case 0: BZErrorCodeToString = "BZ_OK"
        Case 1: BZErrorCodeToString = "BZ_RUN_OK"
        Case 2: BZErrorCodeToString = "BZ_FLUSH_OK"
        Case 3: BZErrorCodeToString = "BZ_FINISH_OK"
        Case 4: BZErrorCodeToString = "BZ_STREAM_END"
        Case (-1): BZErrorCodeToString = "BZ_SEQUENCE_ERROR"
        Case (-2): BZErrorCodeToString = "BZ_PARAM_ERROR"
        Case (-3): BZErrorCodeToString = "BZ_MEM_ERROR"
        Case (-4): BZErrorCodeToString = "BZ_DATA_ERROR"
        Case (-5): BZErrorCodeToString = "BZ_DATA_ERROR_MAGIC"
        Case (-6): BZErrorCodeToString = "BZ_IO_ERROR"
        Case (-7): BZErrorCodeToString = "BZ_UNEXPECTED_EOF"
        Case (-8): BZErrorCodeToString = "BZ_OUTBUFF_FULL"
        Case (-9): BZErrorCodeToString = "BZ_CONFIG_ERROR"
        Case Default: BZErrorCodeToString = "Unknown: " & ec
    End Select
    
End Function

Public Function CompressData(bin() As Byte, bout() As Byte, Optional compLevel As eCompressionLevels = clBest) As Boolean
    Reset
    CompressData = internal_CompressData(bin, bout, compLevel)
End Function

Private Function internal_CompressData(bin() As Byte, bout() As Byte, Optional compLevel As eCompressionLevels = clBest) As Boolean
    
    On Error Resume Next
    Dim ret As Long
    
    If m_lDecompressedSize = 0 Then m_lDecompressedSize = UBound(bin) + 1
    m_lCompressedSize = m_lDecompressedSize + (m_lDecompressedSize * 0.01) + (100 * compLevel)
    
    Erase bout
    ReDim bout(m_lCompressedSize)

    ret = z2CompressAlt(bout(0), m_lCompressedSize, bin(0), m_lDecompressedSize, compLevel, 0, 0)
    
'    If Err.Number Then
'        'win98 support
'        Err.Clear
'        ret = z2Compress(bout(0), m_lCompressedSize, bin(0), m_lDecompressedSize, compLevel, 0, 0)
'        If Err.Number Then Err.Clear
'    End If
     
    If ret = 0 Then
        'NO ERRORS OCCURED DURING COMPRESSION
        ReDim Preserve bout(m_lCompressedSize)
        internal_CompressData = True
    Else
        LastError = "Error compressing data: " & BZErrorCodeToString(ret)
    End If
    
End Function

Public Function CompressFile(infile As String, outfile As String, Optional compLevel As eCompressionLevels = clBest) As Boolean
    Dim bin() As Byte, bout() As Byte
    Dim ret As Long, compressionLevel As Long
    
    Reset
    
    If Not OpenFile(infile, bin, compressionLevel) Then Exit Function
    
    If compressionLevel <> 0 Then
        LastError = "File is already compressed"
        Exit Function
    End If

    If Not internal_CompressData(bin, bout, compLevel) Then Exit Function
    If Not WriteToFile(outfile, bout) Then Exit Function
    CompressFile = True
    
End Function

Public Function DecompressData(bin() As Byte, bout() As Byte) As Boolean
    Reset
    DecompressData = internal_DecompressData(bin, bout)
End Function

Private Function internal_DecompressData(bin() As Byte, bout() As Byte) As Boolean
  Dim lVerbosity As Long ' We want the DLL to shut up, so set it to 0
  Dim lSmall As Long ' if <> 0 then use (s)low memory routines
  Dim ret As Long

    lVerbosity = 0
    lSmall = 0
    
    If m_lCompressedSize = 0 Then m_lCompressedSize = UBound(bin) + 1
    
    If m_lDecompressedSize = 0 Then
        'UNLESS IT WAS IN THE HEADER OF A FILE
        'THERE IS NO WAY TO GUESS THE ORIGINAL SIZE
        m_lDecompressedSize = m_lCompressedSize * 1000
    End If
    
    ReDim bout(m_lDecompressedSize)
  
    'DECOMPRESS
    On Error Resume Next
    ret = z2DecompressAlt(bout(0), m_lDecompressedSize, bin(0), m_lCompressedSize, lSmall, lVerbosity)
    
'    If Err.Number Then
'        'win98 support
'        Err.Clear
'        ret = z2Decompress(bout(0), m_lDecompressedSize, bin(0), m_lCompressedSize, lSmall, lVerbosity)
'    End If
    
    If Err.Number Then
        'ERROR OCCURED
        Err.Clear
    Else
        'NO ERROR OCCURED IN DECOMPRESSION
        If ret = 0 Then
            ReDim Preserve bout(m_lDecompressedSize - 1)
            internal_DecompressData = True
        Else
            LastError = "Error decompressing: " & BZErrorCodeToString(ret)
        End If
    End If
    

End Function


Public Function DecompressFile(infile As String, outfile As String) As Boolean

    Dim bin() As Byte, bout() As Byte, compressionLevel As Long
    
    Reset
    If Not OpenFile(infile, bin, compressionLevel) Then Exit Function
        
    If compressionLevel = 0 Then
        LastError = "File is not compressed"
        Exit Function
    End If
    
    If Not internal_DecompressData(bin, bout) Then Exit Function
    If Not WriteToFile(outfile, bout) Then Exit Function
        
    DecompressFile = True

End Function

'Fills the input buffer: m_abyInputBuffer
Private Function OpenFile(sFilePath, bin() As Byte, Optional ByRef compLevel As Long) As Boolean
 Dim iFileHandle As Integer
 Dim sCheck As String * 3
 Dim sBlockSize As String * 1
 
 On Error GoTo OpenFile_Exit
 
    If Not FileExists(sFilePath) Then
        LastError = "File does not exist: " & sFilePath
        Exit Function
    End If

    'OPEN THE FILE
    compLevel = 0
    iFileHandle = FreeFile
    Open sFilePath For Binary As iFileHandle Len = 32767
    
    'PREPARE THE BUFFER
    m_lCompressedSize = LOF(iFileHandle)
    m_lDecompressedSize = m_lCompressedSize
    ReDim bin(m_lCompressedSize)
    
    'DETERMINE IF THE HEADER HAS THE WORD BZh IN IT
    Get #iFileHandle, , sCheck

    If sCheck = BZh Then
        Get #iFileHandle, , sBlockSize
        compLevel = CByte(sBlockSize)
        
        Get #iFileHandle, , m_lDecompressedSize
        m_lDecompressedSize = m_lDecompressedSize / 1000
    End If
    
    'FILL THE BUFFER
    Seek #iFileHandle, 1
    Get #iFileHandle, , bin()
    OpenFile = True
    
        
OpenFile_Exit:
    Close iFileHandle
    
    If Err.Number <> 0 Then
        LastError = "Error in openFile: " & Err.Description
    End If
    
End Function

Private Sub Reset()
    m_lDecompressedSize = 0
    m_lCompressedSize = 0
    LastError = Empty
End Sub

Private Function WriteToFile(fpath As String, b() As Byte) As Boolean
 Dim fOUT As Integer
 
    On Error GoTo WriteToFile_Exit
 
    If FileExists(fpath) Then Kill fpath

    fOUT = FreeFile
    Open fpath For Binary As fOUT Len = 32767
    Put #fOUT, , b()
    
    WriteToFile = True
    
WriteToFile_Exit:
    On Error Resume Next
    Close fOUT
End Function

'Public Function DirectoryExists(ByVal sPath As String) As Boolean
'    If LenB(sPath) Then
'        If PathIsDirectory(sPath) Then DirectoryExists = True
'    End If
'End Function

Public Function FileExists(ByVal sPath As String) As Boolean
    If LenB(sPath) Then
        If PathFileExists(sPath) Then FileExists = True
    End If
End Function

'Function GetFolder(ByVal sPath As String) As String
'    GetFolder = Left$(sPath, InStrRev(sPath, "\"))
'End Function
