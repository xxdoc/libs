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
'ZipFile
'
'Used by Zipper class to store tasks in a collection of files to Zip.
'

Public SourceFile As String
Public AsFile As String
Public ZMethod As Z_METHODS
Public ZLevel As Z_LEVELS
Public Attrs As VbFileAttribute
Public Comment As String
Public ByteCount As Long
