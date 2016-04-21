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

'4.20.2016 - altered by david zimmer <dzzie@yahoo.com> - simplified

'20031124 altered by mario sepulveda and posted at vbforums.com
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
Private Const BZ_OK As Long = 0
Private Const BZ_RUN_OK As Long = 1
Private Const BZ_FLUSH_OK As Long = 2
Private Const BZ_FINISH_OK As Long = 3
Private Const BZ_STREAM_END As Long = 4
Private Const BZ_SEQUENCE_ERROR As Long = (-1)
Private Const BZ_PARAM_ERROR As Long = (-2)
Private Const BZ_MEM_ERROR As Long = (-3)
Private Const BZ_DATA_ERROR As Long = (-4)
Private Const BZ_DATA_ERROR_MAGIC As Long = (-5)
Private Const BZ_IO_ERROR As Long = (-6)
Private Const BZ_UNEXPECTED_EOF As Long = (-7)
Private Const BZ_OUTBUFF_FULL As Long = (-8)
Private Const BZ_CONFIG_ERROR As Long = (-9)

Private Const BZh As String * 3 = "BZh"

'Declares
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Private Declare Function z2Compress Lib "libbz2.dll" Alias "bzBuffToBuffCompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal blockSize100k As Long, ByVal Verbosity As Long, ByVal workFactor As Long) As Long
Private Declare Function z2Decompress Lib "libbz2.dll" Alias "bzBuffToBuffDecompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal Small As Long, ByVal Verbosity As Long) As Long
Private Declare Function z2CompressAlt Lib "libbz2.dll" Alias "BZ2_bzBuffToBuffCompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal blockSize100k As Long, ByVal Verbosity As Long, ByVal workFactor As Long) As Long
Private Declare Function z2DecompressAlt Lib "libbz2.dll" Alias "BZ2_bzBuffToBuffDecompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal Small As Long, ByVal Verbosity As Long) As Long

'Enums
Public Enum eCompressionLevels
    cl_Fast = 1
    cl_1 = 1
    cl_2 = 2
    cl_3 = 3
    cl_4 = 4
    cl_5 = 5
    cl_6 = 6
    cl_7 = 7
    cl_8 = 8
    cl_9 = 9
    cl_Best = 9
End Enum

'Properties
 Private m_eCompressionLevel As eCompressionLevels
 Private m_abyInputBuffer() As Byte
 Private m_abyOutputBuffer() As Byte
 Private m_sInputFilePath As String
 Private m_sOutputFilePath As String
 Private m_bOverwrite As Boolean
 Private m_bKeep As Boolean
 Private m_lDecompressedSize As Long '-In Bytes
 Private m_lCompressedSize As Long   '-In Bytes
 Private m_bClearBuffer As Boolean
 Private m_bIsCompressed As Boolean
 

Public Function CompressData(bin() As Byte, bout() As Byte) As Boolean
    Reset
    CompressData = internal_CompressData(bin, bout)
End Function

Private Function internal_CompressData(bin() As Byte, bout() As Byte) As Boolean
    
    Dim ret As Long
    
    If m_lDecompressedSize = 0 Then m_lDecompressedSize = UBound(bin) + 1
    m_lCompressedSize = m_lDecompressedSize + (m_lDecompressedSize * 0.01) + (100 * m_eCompressionLevel)
    
    Erase bout
    ReDim bout(m_lCompressedSize)
  
 On Error Resume Next
 
    ret = z2CompressAlt(bout(0), m_lCompressedSize, bin(0), m_lDecompressedSize, m_eCompressionLevel, 0, 0)
    
    If Err.Number Then
        Err.Clear
        ret = z2Compress(bout(0), m_lCompressedSize, bin(0), m_lDecompressedSize, m_eCompressionLevel, 0, 0)
        If Err.Number Then Err.Clear
    End If
    
 On Error GoTo 0
    
    If ret = 0 Then
        'NO ERRORS OCCURED DURING COMPRESSION
        ReDim Preserve bout(m_lCompressedSize)
        internal_CompressData = True
    End If
    
End Function

Public Function CompressFile(infile As String, outfile As String) As Boolean
    
    Reset
    m_sInputFilePath = infile
    m_sOutputFilePath = outfile
    
    Dim ret As Long
    
    If OpenFile(infile) Then
        If m_bIsCompressed Then
            'ALREADY COMPRESSED WITH BZIP2
            'CompressFile = 2
        Else
            'THIS IS WHERE THE DATA IS COMPRESSED
            If internal_CompressData(m_abyInputBuffer, m_abyOutputBuffer) Then
                If WriteToFile Then CompressFile = True
            End If
        End If
    End If
    
    
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
    
    If Err.Number Then
        'ERROR OCCURED
        Err.Clear
        ret = z2Decompress(bout(0), m_lDecompressedSize, bin(0), m_lCompressedSize, lSmall, lVerbosity)
    End If
    
    If Err.Number Then
        'ERROR OCCURED
        Err.Clear
    Else
        'NO ERROR OCCURED IN DECOMPRESSION
        If ret = 0 Then
            ReDim Preserve bout(m_lDecompressedSize - 1)
            internal_DecompressData = True
        End If
        
    End If
 On Error GoTo 0
End Function

'opens a file
'fills the buffer
'decompresses data
'writes the decompressed file
'returns 0 if succesfull
Public Function DecompressFile(infile As String, outfile As String) As Boolean

    Reset
    m_sInputFilePath = infile
    m_sOutputFilePath = outfile

    If OpenFile(infile) Then
        If LenB(m_sOutputFilePath) Then
            'OUTPUT PATH HAS BEEN SPECIFIED
            If m_bIsCompressed Then
                'THIS WHERE THE DATA IS ACTUALLY COMPRESSED
                If internal_DecompressData(m_abyInputBuffer, m_abyOutputBuffer) Then
                    'NO ERRORS DECOMPRESSING
                    If WriteToFile Then DecompressFile = True
                End If
            Else
                'THIS FILE WAS NOT COMPRESSED WITH BZIP2
                'DecompressFile = 9
            End If
        End If
    End If
    
End Function

Public Function DirectoryExists(ByVal sPath As String) As Boolean
    If LenB(sPath) Then
        If PathIsDirectory(sPath) Then DirectoryExists = True
    End If
End Function

Public Function FileExists(ByVal sPath As String) As Boolean
    If LenB(sPath) Then
        If PathFileExists(sPath) Then FileExists = True
    End If
End Function

Friend Function GetFolder(ByVal sPath As String) As String
    GetFolder = Left$(sPath, InStrRev(sPath, "\"))
End Function

'Opens a file
'Fills the input buffer: m_abyInputBuffer
Private Function OpenFile(Optional ByVal sFilePath As String = vbNullString) As Boolean
 Dim iFileHandle As Integer
 Dim sCheck As String * 3
 Dim sBlockSize As String * 1
 Dim bContinue As Boolean
 
 On Error GoTo OpenFile_Exit
    If FileExists(sFilePath) Then
        m_sInputFilePath = sFilePath
        bContinue = True
    Else
        If FileExists(m_sInputFilePath) Then bContinue = True
    End If
    
    If bContinue Then
        'FILE EXISTS
        Erase m_abyInputBuffer
        Erase m_abyOutputBuffer
    
        'OPEN THE FILE
        iFileHandle = FreeFile
        Open m_sInputFilePath For Binary As iFileHandle Len = 32767
        
        'PREPARE THE BUFFER
        m_lCompressedSize = LOF(iFileHandle)
        ReDim m_abyInputBuffer(m_lCompressedSize)
        
        'DETERMINE IF THE HEADER HAS THE WORD "BZh IN IT
        Get #iFileHandle, , sCheck
        m_bIsCompressed = (sCheck = BZh)
    
        If m_bIsCompressed Then
            'ALREADY COMPRESSED WITH BZIP2
            If LenB(m_sOutputFilePath) = 0 Then
                'm_sOutputFilePath is empty
                If Right$(m_sInputFilePath, 4) = ".bz2" Then
                    m_sOutputFilePath = Left$(m_sInputFilePath, Len(m_sInputFilePath) - 4)
                    If Len(m_sOutputFilePath) <= 3 Then m_sOutputFilePath = vbNullString 'not a real path
                End If
            End If
                    
            Get #iFileHandle, , sBlockSize
            m_eCompressionLevel = CByte(sBlockSize)
            
            Get #iFileHandle, , m_lDecompressedSize
            m_lDecompressedSize = m_lDecompressedSize / 1000
        Else
            'NOT COMPRESSED YET
            If LenB(m_sOutputFilePath) = 0 Then m_sOutputFilePath = m_sInputFilePath & ".bz2"
        
            m_lDecompressedSize = m_lCompressedSize
        End If
        
        'FILL THE BUFFER
        Seek #iFileHandle, 1
        Get #iFileHandle, , m_abyInputBuffer
        OpenFile = True
    Else
        Err.Raise 75
    End If
        
OpenFile_Exit:
    Close iFileHandle
End Function

Private Sub Reset()
    m_eCompressionLevel = cl_Best
    Erase m_abyInputBuffer
    Erase m_abyOutputBuffer
    m_sInputFilePath = vbNullString
    m_sOutputFilePath = vbNullString
    m_bOverwrite = True
    m_bKeep = False
    m_lDecompressedSize = 0
    m_lCompressedSize = 0
    m_bClearBuffer = True
    m_bIsCompressed = False
End Sub

Private Function WriteToFile() As Boolean
 Dim fOUT As Integer
 
 
    On Error GoTo WriteToFile_Exit
 
    If FileExists(m_sOutputFilePath) Then Kill m_sOutputFilePath

    fOUT = FreeFile
    Open m_sOutputFilePath For Binary As fOUT Len = 32767
    Put #fOUT, , m_abyOutputBuffer
    
    If m_bClearBuffer Then Erase m_abyOutputBuffer
    WriteToFile = True
    
WriteToFile_Exit:
    On Error Resume Next
    Close fOUT
End Function
'
'Public Property Get InputFileIsCompessed() As Boolean
'    InputFileIsCompessed = m_bIsCompressed
'End Property
