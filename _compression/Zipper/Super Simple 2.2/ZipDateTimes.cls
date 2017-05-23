VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipDateTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Some useful notes on filesystem timestamps that the MSDN Library seems to
'have recently lost:
'
'===========================================
'WIN32_FILE_ATTRIBUTE_DATA Structure REMARKS
'===========================================
'
'Not all file systems can record creation and last access time, and not all
'file systems record them in the same manner. For example, on the FAT file
'system, create time has a resolution of 10 milliseconds, write time has a
'resolution of 2 seconds, and access time has a resolution of 1 day. On the
'NTFS file system, access time has a resolution of 1 hour. For more
'information, see File Times.
'
'==========
'File Times
'==========
'
'A file time is a 64-bit value that represents the number of 100-nanosecond
'intervals that have elapsed since 12:00 A.M. January 1, 1601 (UTC). The
'system records file times whenever applications create, access, and write
'to files. Not all file systems can record creation and last access time and
'not all file systems record them in the same manner. For example, on NT FAT,
'create time has a resolution of 10 milliseconds, write time has a resolution
'of 2 seconds, and access time has a resolution of 1 day (really, the access
'date). On NTFS, access time has a resolution of 1 hour.
'
'To retrieve the file times for a specified file, use the GetFileTime function.
'GetFileTime copies the creation, last access, and last write times to
'individual FILETIME structures. You can also retrieve file times using the
'FindFirstFile and FindNextFile functions. These functions copy the file times
'to FILETIME structures in a WIN32_FIND_DATA structure.
'
'To set the file times for a file, use the SetFileTime function. This function
'lets you modify creation, last access, and last write times without changing
'the content of the file. You can compare the times of different files by
'using the CompareFileTime function. The function compares two file times and
'returns a value indicating which time is greater or returns zero if the times
'are equal.
'
'If you plan to modify file times for specified files, you can convert a date
'and time of day to a file time by using the SystemTimeToFileTime function.
'You can also obtain the system time in a FILETIME structure by calling the
'GetSystemTimeAsFileTime function.
'
'To make a file time easy to display to a user, use the FileTimeToSystemTime
'function. FileTimeToSystemTime converts the file time and copies the month,
'day, year, and time of day from the file time to a SYSTEMTIME structure.
'
'===================================
'File Times and Daylight Saving Time
'===================================
'
'You must take care using file times if the user has set the system to
'automatically adjust for daylight saving time.
'
'To convert a file time to local time, use the FileTimeToLocalFileTime
'function. However, FileTimeToLocalFileTime uses the current settings for the
'time zone and daylight saving time. Therefore, if it is daylight saving time,
'it will take daylight saving time into account, even if the file time you are
'converting is in standard time.
'
'FAT records times on disk in local time. GetFileTime retrieves cached UTC
'times from FAT. When it becomes daylight saving time, the time retrieved by
'GetFileTime will be off an hour, because the cache has not been updated. When
'you restart the machine, the cached time retrieved by GetFileTime will be
'correct. FindFirstFile retrieves the local time from FAT and converts it to
'UTC using the current settings for the time zone and daylight saving time.
'Therefore, if it is daylight saving time, FindFirstFile will take daylight
'saving time into account, even if the file time you are converting is in
'standard time.
'
'NTFS records times on disk in UTC. To account for daylight saving time when
'converting a file time to a local time, use the following functions instead
'of FileTimeToLocalFileTime:
'
'       FileTimeToSystemTime
'       SystemTimeToTzSpecificLocalTime
'       SystemTimeToFileTime

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FILE_ATTRIBUTE_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
End Type
    
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
    
Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    ByRef lpFileTime As FILETIME, _
    ByRef lpSystemTime As SYSTEMTIME) As Long

Private Declare Function GetFileAttributesEx Lib "kernel32" _
    Alias "GetFileAttributesExW" ( _
    ByVal lpFileName As Long, _
    ByVal fInfoLevelId As Long, _
    ByVal lpFileInformation As Long) As Long

Private Declare Function GetTimeZoneInformation Lib "kernel32" ( _
    ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" ( _
    ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
    ByRef lpUniversalTime As SYSTEMTIME, _
    ByRef lpLocalTime As SYSTEMTIME) As Long

Private Const GetFileExInfoStandard As Long = 0

'These are assigned Date values (or #12:00:00 AM# if not available).
Private mCreated As Date
Private mLastAccessed As Date
Private mLastModified As Date

Public Property Get Created() As Date
    Created = mCreated
End Property

Public Property Get LastAccessed() As Date
    LastAccessed = mLastAccessed
End Property

Public Property Get LastModified() As Date
    LastModified = mLastModified
End Property

Public Sub GetTimestamps(ByVal PathToCheck As String)
    'Retrieve a subset of filesystem info for the file or directory in
    'PathToCheck and cache results for Property-based retrieval.
    '
    'We are assuming NTFS here.  Objects on a FAT partition may give
    'times off by an hour.  See the lengthy notes above.
    '
    'Non-existant timestamps (not supported by the filesystem) return
    'Empty values.
    
    Dim wfadFile As WIN32_FILE_ATTRIBUTE_DATA
    Dim lngResult As Long

    lngResult = GetFileAttributesEx(StrPtr(PathToCheck), _
                                    GetFileExInfoStandard, _
                                    VarPtr(wfadFile))
    If lngResult = 0 Then
        Err.Raise &H8004AA51, _
                  TypeName(Me), _
                  "Internal error. System error " _
                & CStr(Err.LastDllError) _
                & " calling GetFileAtributesEx"
    Else
        mCreated = FileTimeToDate(wfadFile.ftCreationTime)
        mLastAccessed = FileTimeToDate(wfadFile.ftLastAccessTime)
        mLastModified = FileTimeToDate(wfadFile.ftLastWriteTime)
    End If
End Sub

Private Sub UniversalToLocal(ByRef UniversalTime As SYSTEMTIME, ByRef LocalTime As SYSTEMTIME)
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION

    GetTimeZoneInformation TimeZoneInfo
    If SystemTimeToTzSpecificLocalTime(TimeZoneInfo, UniversalTime, LocalTime) = 0 Then
        Err.Raise &H8004AA53, _
                  TypeName(Me), _
                  "Internal error. System error " _
                & CStr(Err.LastDllError) _
                & " calling SystemTimeToTzSpecificLocalTime"
    End If
End Sub

Private Function FileTimeToDate(ByRef FILETIME As FILETIME) As Date
    Dim stUniversal As SYSTEMTIME
    Dim stLocal As SYSTEMTIME
    
    With FILETIME
        If (.dwHighDateTime <> 0) And (.dwLowDateTime <> 0) Then
            FileTimeToSystemTime FILETIME, stUniversal
            UniversalToLocal stUniversal, stLocal
            With stLocal
                FileTimeToDate = DateSerial(.wYear, .wMonth, .wDay) _
                               + TimeSerial(.wHour, .wMinute, .wSecond)
            End With
        'Else
            'FileTimeToDate = #12:00:00 AM#
        End If
    End With
End Function
