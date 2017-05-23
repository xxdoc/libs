VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'ZipFile version 2.2
'
'Used by Zipper to store tasks in a collection of files to Zip.
'
'Changes:
'
'   Version 2:
'
'       * Adds password encyption.
'
'       * Adds file date specification.
'
'   Version 2.1:
'
'       No change.
'
'   Version 2.2:
'
'       * Added BLOB support.  If Len(SourceFile) = 0 then BLOB
'         is the data to add instead of a file.  At this writing
'         only ZipperSync implemented this functionality, though
'         it could also be done in the Zipper control with some
'         changes.
'
'         Note that since array-valued properties don't work, we
'         implement this as GetBLOB()/LetBLOB() here.
'

Private mBLOB() As Byte

Public SourceFile As String
Public AsFile As String
Public ZMethod As Z_METHODS
Public ZLevel As Z_LEVELS
Public Attrs As VbFileAttribute
Public Comment As String
Public FileDate As Date
Public Password As String
Public NeedCRC As Boolean 'Should be set to True when Password is used (i.e. not an empty
                          'string).  Gets set back to False after CRC calculation for pass 2
                          '(when opening and writing to the ZIP archive occurs).
Public ByteCount As Long

Public Function GetBLOB() As Byte()
    GetBLOB = mBLOB
End Function

Public Sub LetBLOB(ByRef RHS() As Byte)
    mBLOB = RHS
End Sub

Private Sub Class_Initialize()
    mBLOB = "" 'Empty Byte array.
End Sub
