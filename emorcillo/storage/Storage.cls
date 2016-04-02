VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Storage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Storage class
'
' IStorage wrapper class
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 08/03/1999
' Updates: 02/17/2000 * Added Parent property. This property
'                       will keep the parent storage alive
'                       if you use the same object for all
'                       storages in your file.
'
'*********************************************************************************************
Option Explicit

Private m_oStorage As olelib.IStorage
Private m_oStat As StatStg
Private m_oParent As Storage

Enum MoveFlags
   sfCopy = STGMOVE.STGMOVE_COPY
   sfMove = STGMOVE.STGMOVE_MOVE
   sfShallowCopy = STGMOVE.STGMOVE_SHALLOWCOPY
End Enum

'*********************************************************************************************
'
' Creates the standard property sets
'
' Parameters:
'
' oStg   - target storage
' ANSI   - If True the properties will use ANSI string
'
'*********************************************************************************************
Public Sub CreatePropertySets(Optional ByVal ANSI As Boolean = True)
Dim oPSS As IPropertySetStorage
Dim tFMTID As olelib.UUID
Dim tCLSID As olelib.UUID
Dim lFlags As Long

   On Error Resume Next

   ' Get IPropertySetStorage interface
   Set oPSS = m_oStorage

   ' Set flags
   If ANSI Then
       lFlags = PROPSETFLAG_ANSI
   Else
       lFlags = PROPSETFLAG_DEFAULT
   End If

   ' Create SummaryInformation, DocumentSummaryInformation
   ' and UserProperties property set storages
   oPSS.Create FMTID_SummaryInformation, tCLSID, lFlags, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE
   oPSS.Create FMTID_DocSummaryInformation, tCLSID, lFlags, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE
   oPSS.Create FMTID_UserProperties, tCLSID, lFlags, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE

End Sub

'*********************************************************************************************
' Commit: Save all changes to disk
'*********************************************************************************************
Public Sub Commit(Optional ByVal Flags As STGC = STGC_DEFAULT)
   On Error GoTo ShowError

   m_oStorage.Commit Flags

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
' CreateStorage: Creates a storage object within this storage
'*********************************************************************************************
Public Function CreateStorage(ByVal Name As String, Optional ByVal Flags As OpenFlags = sfReadWrite Or sfShareExclusive) As Storage
Dim IStg As olelib.IStorage

   On Error GoTo ShowError

   ' Create the olelib.IStorage object
   Set IStg = m_oStorage.CreateStorage(Name, Flags, 0)

   ' Create a new m_oStorage object
   Set CreateStorage = New Storage
   CreateStorage.frInitialize IStg, Me

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
' CreateStream: Creates a new stream within this storage
'*********************************************************************************************
Public Function CreateStream(ByVal Name As String, Optional ByVal Flags As OpenFlags = sfReadWrite Or sfShareExclusive) As Stream
Dim IStrm As olelib.IStream

   On Error GoTo ShowError

   ' Create the olelib.IStream object
   Set IStrm = m_oStorage.CreateStream(Name, Flags)

   ' Create the stream object
   Set CreateStream = New Stream

   ' Initialize the object
   CreateStream.frInitialize IStrm, Me

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
' DestroyElement: Removes a stream or storage from this storage
'*********************************************************************************************
Public Sub DestroyElement(ByVal Name As String)

   On Error GoTo ShowError

   m_oStorage.DestroyElement Name

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
' Elements: Returns a Elements collection with all the storage elements
'*********************************************************************************************
Public Function Elements() As Elements
Dim oEnm As IEnumSTATSTG
Dim tSSTG As olelib.StatStg

   On Error GoTo ShowError

   ' Get the storage enumerator object
   Set oEnm = m_oStorage.EnumElements()

   ' Create a new Elements Object
   Set Elements = New Elements

   ' Enumerate all elements and
   ' add them to the collection
   Do While oEnm.Next(1, tSSTG) = S_OK
       Elements.Add tSSTG
   Loop

   Set oEnm = Nothing

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
' MoveElementTo: Moves or copies elements from this storage to another
'*********************************************************************************************
Public Sub MoveElementTo(ByVal ElementName As String, ByVal DestStrg As Storage, ByVal NewName As String, Optional ByVal Flags As MoveFlags = sfMove)

   On Error GoTo ShowError

   m_oStorage.MoveElementTo ElementName, DestStrg.Storage, NewName, Flags

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
' OpenStorage: Opens a storage within this storage
'*********************************************************************************************
Public Function OpenStorage(ByVal Name As String, Optional ByVal Flags As OpenFlags = sfReadWrite Or sfShareExclusive) As Storage
Dim oStg As olelib.IStorage

   ' Open the storage
   Set oStg = m_oStorage.OpenStorage(Name, 0, Flags)

   ' Create a new Storage object
   Set OpenStorage = New Storage

   ' Initialize the object
   OpenStorage.frInitialize oStg, Me

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
' OpenStream: opens a stream within this storage
'*********************************************************************************************
Public Function OpenStream(ByVal Name As String, Optional ByVal Flags As STGM = STGM_READWRITE Or STGM_SHARE_EXCLUSIVE) As Stream
Dim oStrm As olelib.IStream

   On Error GoTo ShowError

   ' Open the stream
   Set oStrm = m_oStorage.OpenStream(Name, 0, Flags)

   ' Create a new stream object
   Set OpenStream = New Stream
   OpenStream.frInitialize oStrm, Me

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

'
' Parent: Returns the parent storage or Nothing if this is the root storage
'
Public Property Get Parent() As Storage
   Set Parent = m_oParent
End Property

'*********************************************************************************************
' RenameElement: Renames a stream or storage within this storage
'*********************************************************************************************
Public Sub RenameElement(ByVal OldName As String, NewName As String)

   On Error GoTo ShowError

   m_oStorage.RenameElement OldName, NewName

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
' Revert: Reverts all non comitted changes
'*********************************************************************************************
Public Sub Revert()

   On Error GoTo ShowError

   m_oStorage.Revert

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
' Initialize: Initializes the object
'*********************************************************************************************
Friend Sub frInitialize(oStg As olelib.IStorage, oParent As Storage)
Dim SSTG As olelib.StatStg

   ' Set the storage
   Set m_oStorage = oStg

   ' Set the parent
   Set m_oParent = oParent

   ' Get m_oStorage info
   m_oStorage.Stat SSTG

   ' Create new STATSTG object
   Set m_oStat = New StatStg
   m_oStat.frInitialize SSTG

End Sub

Public Property Get Storage() As olelib.IStorage
    Set Storage = m_oStorage
End Property

'*********************************************************************************************
' Stat: returns a STATSTG object with info about this storage
'*********************************************************************************************
Public Property Get Stat() As StatStg
    Set Stat = m_oStat
End Property

Private Sub Class_Terminate()

   On Error Resume Next

   m_oStorage.Commit STGC_OVERWRITE

End Sub

