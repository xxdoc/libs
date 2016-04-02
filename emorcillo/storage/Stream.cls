VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Stream class
'
' IStream wrapper class
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 08/03/1999
' Updates:
'          08/03/1999. * WriteData was renamed to WriteBuf.
'                      * New WriteData with Buf as Variant.
'                      * ReadData was renamed to ReadBuf
'                      * New ReadData with Buf as Variant.
'
'          12/13/1999. * Stream public variable moved to
'                        read-only property
'                      * New method WriteObject
'                      * New method ReadObject
'
'          02/17/2000. * Added Parent property
'*********************************************************************************************
Option Explicit

Private m_oStream As olelib.IStream
Private m_oStat As STATSTG
Private m_oParent As Storage

'*********************************************************************************************
' Clone: Duplicates this object
'*********************************************************************************************
Public Function Clone() As Stream

   On Error GoTo ShowError
   
   Set Clone = New Stream
   
   Clone.frInitialize m_oStream.Clone, m_oParent
    
   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If
   
End Function

'*********************************************************************************************
' Commit: Save all changes to disk
'*********************************************************************************************
Public Sub Commit(Optional ByVal Flags As STGC = STGC_DEFAULT)

   On Error GoTo ShowError
   
   m_oStream.Commit Flags
    
   Exit Sub
    
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If

End Sub

'*********************************************************************************************
' CopyTo: Copy BufLen bytes from this stream to another
'*********************************************************************************************
Public Function CopyTo(ByVal Strm As Stream, ByVal BufLen As Currency) As Currency
Dim Rd As Currency

   On Error GoTo ShowError
   
   BufLen = BufLen / 10000
   
   m_oStream.CopyTo Strm.Stream, BufLen, Rd, CopyTo
   
   CopyTo = CopyTo * 10000
   
   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If
    
End Function

Friend Sub frInitialize(Strm As olelib.IStream, Parent As Storage)
Dim SSTG As olelib.STATSTG

   ' Set the storage
   Set m_oStream = Strm

   ' Set the parent
   Set m_oParent = Parent

   ' Get m_oStorage info
   m_oStream.Stat SSTG

   ' Create new STATSTG object
   Set m_oStat = New STATSTG
   m_oStat.frInitialize SSTG

End Sub

'*********************************************************************************************
' Parent: returns the parent storage
'*********************************************************************************************
Public Property Get Parent() As Storage
   Set Parent = m_oParent
End Property

'*********************************************************************************************
' Revert: reverts all non comitted changes
'*********************************************************************************************
Public Sub Revert()
    
   On Error GoTo ShowError
   
   m_oStream.Revert
    
   Exit Sub
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If

End Sub

'*********************************************************************************************
' SetSize: Sets the stream size
'*********************************************************************************************
Public Sub SetSize(ByVal NewSize As Currency)
    
   On Error GoTo ShowError
    
   m_oStream.SetSize NewSize / 10000

   Exit Sub
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If

End Sub

'*********************************************************************************************
' ReadBuf: reads BufLen bytes from the stream and stores
'          them in the Buf array.
'*********************************************************************************************
Public Function ReadBuf(Buf() As Byte, ByVal BufLen As Long) As Long

   On Error GoTo ShowError

   ReDim Buf(0 To BufLen - 1)
    
   ReadBuf = m_oStream.Read(Buf(0), BufLen)

   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If
    
End Function

'*********************************************************************************************
' ReadObject: Reads and creates an object previously saved with WriteObject
'*********************************************************************************************
Public Function ReadObject() As Object
   
   On Error GoTo ShowError
   
   Set ReadObject = mdlStreamFuncs.ReadObject(m_oStream)
   
   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum, , Err.Description
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If
   
End Function

'*********************************************************************************************
' ReadData: Reads data in the VType format and returns
'           it as a Variant
'*********************************************************************************************
Public Function ReadData(ByVal VType As VbVarType) As Variant

   On Error GoTo ShowError
   
   ReadData = ReadValue(m_oStream, VType)
   
   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If
    
End Function
'*********************************************************************************************
' SSeek: Moves the current stream position to a new one
' (this method is called SSeek because Seek is a reserved word)
'*********************************************************************************************
Public Function SSeek(ByVal NewPos As Currency) As Currency

    NewPos = NewPos / 10000
    
    SSeek = m_oStream.Seek(NewPos, 0) * 10000
    
End Function

'*********************************************************************************************
' Stat: Returns a STATSTG with stream info
'*********************************************************************************************
Public Function Stat() As STATSTG
    
    Set Stat = m_oStat
    
End Function

'*********************************************************************************************
' Stream: Returns/Sets(within the same project) the olelib.IStream interface
'*********************************************************************************************
Public Property Get Stream() As olelib.IStream
    
    Set Stream = m_oStream
    
End Property

'*********************************************************************************************
' WriteBuff: Writes an array in the stream
'*********************************************************************************************
Public Function WriteBuf(Buf() As Byte) As Long

   On Error GoTo ShowError
   
   WriteBuf = m_oStream.Write(Buf(LBound(Buf)), UBound(Buf) - LBound(Buf) + 1)
    
   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If
    
End Function

'*********************************************************************************************
' WriteData: Writes the data stored in a Variant. Use this method to save
'            directly from variables.
'*********************************************************************************************
Public Function WriteData(Buf As Variant) As Long
   
   On Error GoTo ShowError
   
   WriteValue m_oStream, Buf, VariantType(Buf)
   
   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If

End Function

'*********************************************************************************************
' WriteObject: Writes an object to the stream. Use this method to save pictures,
'              fonts and any object that implements IPersistStream interface.
'*********************************************************************************************
Public Sub WriteObject(ByVal Obj As Object)
   
   On Error GoTo ShowError
   
   mdlStreamFuncs.WriteObject m_oStream, Obj
   
   Exit Sub
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number

   If ErrNum > 0 Then
      Err.Raise ErrNum
   Else
      Err.Raise ErrNum, , ErrorMessage(ErrNum)
   End If
   
End Sub
