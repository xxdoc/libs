VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatStg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' STATSTG
'
' STATSTG UDT wrapper class
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 08/03/1999
'
'*********************************************************************************************

Option Explicit

Dim m_Stat As olelib.StatStg

Enum ElementType
   sfLockBytes = STGTY_LOCKBYTES
   sfStorage = STGTY_STORAGE
   sfStream = STGTY_STREAM
   sfProperty = STGTY_PROPERTY
End Enum


Public Property Get Mode() As OpenFlags
   Mode = m_Stat.grfMode
End Property

Public Property Get Size() As OpenFlags
   Size = m_Stat.cbSize * 10000
End Property



Public Property Get Name() As String
    Name = Ptr2Str(m_Stat.pwcsName, False, True)
End Property

Public Property Get CLSID() As String
    CLSID = Space$(38)
    olelib.StringFromGUID2 m_Stat.CLSID, CLSID, 39
End Property
Public Property Get ElemType() As ElementType
   ElemType = m_Stat.Type
End Property

Friend Sub frInitialize(SSTG As olelib.StatStg)
    m_Stat = SSTG
End Sub

Private Sub Class_Terminate()
    
    ' Free memory used by the struct
    CoTaskMemFree m_Stat.pwcsName
    
End Sub



