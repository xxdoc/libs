Attribute VB_Name = "AutoCompleteHandle"
Option Explicit

' Required:

' AutoCompleteGuids.tlb (in IDE only)

Private Enum VTableIndexIEnumStringConstants
' Ignore : IEnumStringQueryInterface = 1
' Ignore : IEnumStringAddRef = 2
' Ignore : IEnumStringRelease = 3
VTableIndexIEnumStringNext = 4
VTableIndexIEnumStringSkip = 5
VTableIndexIEnumStringReset = 6
VTableIndexIEnumStringClone = 7
End Enum
Private Type IEnumStringDispatcherStruct
This As AutoCompleteGuids.IEnumStringVB
IEnumStringPtr As Long
End Type
Private Type Guid
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Const CLSCTX_INPROC_SERVER As Long = 1
Private Const CLSID_IAutoComplete As String = "{00BB2763-6A77-11D0-A535-00C04FD7D062}"
Private Const CLSID_IAutoCompleteDropDown As String = "{3CD141F4-3C6A-11D2-BCAA-00C04FD929DB}"
Private Const CLSID_IACList2 As String = "{470141A0-5186-11D2-BBB6-0060977B464C}"
Private Const CLSID_ACLHistory As String = "{00BB2764-6A77-11D0-A535-00C04FD7D062}"
Private Const CLSID_ACListISF As String = "{03C036F1-A186-11D0-824A-00AA005B4383}"
Private Const CLSID_ACLMRU As String = "{6756A641-DE71-11D0-831B-00AA005B4383}"
Private Const CLSID_ACLMulti As String = "{00BB2765-6A77-11D0-A535-00C04FD7D062}"
Private Const IID_IAutoComplete2 As String = "{EAC04BC0-3791-11D2-BB95-0060977B464C}"
Private Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
Private Const IID_IObjMgr As String = "{00BB2761-6A77-11D0-A535-00C04FD7D062}"
Private IEnumStringDispatcher() As IEnumStringDispatcherStruct
Private IEnumStringDispatcherCount As Long
Private VTableSubclassIEnumString As VTableSubclass

Public Sub SetVTableSubclassIEnumString(ByVal This As Object)
If VTableSupported(This) = True Then
    Dim ShadowIEnumString As AutoCompleteGuids.IEnumString
    Set ShadowIEnumString = This
    ReDim Preserve IEnumStringDispatcher(0 To IEnumStringDispatcherCount) As IEnumStringDispatcherStruct
    Set IEnumStringDispatcher(IEnumStringDispatcherCount).This = This
    IEnumStringDispatcher(IEnumStringDispatcherCount).IEnumStringPtr = ObjPtr(ShadowIEnumString)
    IEnumStringDispatcherCount = IEnumStringDispatcherCount + 1
    Call ReplaceIEnumString(This)
End If
End Sub

Public Sub RemoveVTableSubclassIEnumString(ByVal This As Object)
If VTableSupported(This) = True Then Call RestoreIEnumString(This)
End Sub

Private Function VTableSupported(ByRef This As Object) As Boolean
On Error GoTo Cancel
Dim ShadowIEnumString As AutoCompleteGuids.IEnumString
Dim ShadowIEnumStringVB As AutoCompleteGuids.IEnumStringVB
Set ShadowIEnumString = This
Set ShadowIEnumStringVB = This
VTableSupported = Not CBool(ShadowIEnumString Is Nothing Or ShadowIEnumStringVB Is Nothing)
Cancel:
End Function

Private Sub ReplaceIEnumString(ByVal This As AutoCompleteGuids.IEnumString)
If VTableSubclassIEnumString Is Nothing Then Set VTableSubclassIEnumString = New VTableSubclass
If VTableSubclassIEnumString.RefCount = 0 Then
    VTableSubclassIEnumString.Subclass ObjPtr(This), VTableIndexIEnumStringNext, VTableIndexIEnumStringClone, _
    AddressOf IEnumString_Next, AddressOf IEnumString_Skip, _
    AddressOf IEnumString_Reset, AddressOf IEnumString_Clone
End If
VTableSubclassIEnumString.AddRef
End Sub

Private Sub RestoreIEnumString(ByVal This As AutoCompleteGuids.IEnumString)
If Not VTableSubclassIEnumString Is Nothing Then
    VTableSubclassIEnumString.Release
    If VTableSubclassIEnumString.RefCount = 0 Then VTableSubclassIEnumString.UnSubclass
End If
End Sub

Private Function IEnumString_Next(ByVal This As Long, ByVal cElt As Long, ByVal rgElt As Long, ByVal pcEltFetched As Long) As Long
Dim i As Long
For i = 0 To IEnumStringDispatcherCount - 1
    If IEnumStringDispatcher(i).IEnumStringPtr = This Then IEnumStringDispatcher(i).This.Next IEnumString_Next, cElt, rgElt, pcEltFetched
Next i
End Function

Private Function IEnumString_Skip(ByVal This As Long, ByVal cElt As Long) As Long
Dim i As Long
For i = 0 To IEnumStringDispatcherCount - 1
    If IEnumStringDispatcher(i).IEnumStringPtr = This Then IEnumStringDispatcher(i).This.Skip IEnumString_Skip, cElt
Next i
End Function

Private Function IEnumString_Reset(ByVal This As Long) As Long
Dim i As Long
For i = 0 To IEnumStringDispatcherCount - 1
    If IEnumStringDispatcher(i).IEnumStringPtr = This Then IEnumStringDispatcher(i).This.Reset IEnumString_Reset
Next i
End Function

Private Function IEnumString_Clone(ByVal This As Long, ByRef ppEnum As AutoCompleteGuids.IEnumString) As Long
Dim i As Long
For i = 0 To IEnumStringDispatcherCount - 1
    If IEnumStringDispatcher(i).IEnumStringPtr = This Then IEnumStringDispatcher(i).This.Clone IEnumString_Clone, ObjPtr(ppEnum)
Next i
End Function

Public Function CreateIAutoComplete2() As IUnknown
Dim CLSID As Guid, IID As Guid
On Error Resume Next
CLSIDFromString StrPtr(CLSID_IAutoComplete), CLSID
CLSIDFromString StrPtr(IID_IAutoComplete2), IID
CoCreateInstance CLSID, 0, CLSCTX_INPROC_SERVER, IID, CreateIAutoComplete2
End Function

Public Function CreateIAutoCompleteDropDown() As IUnknown
Dim CLSID As Guid, IID As Guid
On Error Resume Next
CLSIDFromString StrPtr(CLSID_IAutoCompleteDropDown), CLSID
CLSIDFromString StrPtr(IID_IUnknown), IID
CoCreateInstance CLSID, 0, CLSCTX_INPROC_SERVER, IID, CreateIAutoCompleteDropDown
End Function

Public Function CreateIACList2() As IUnknown
Dim CLSID As Guid, IID As Guid
On Error Resume Next
CLSIDFromString StrPtr(CLSID_IACList2), CLSID
CLSIDFromString StrPtr(IID_IUnknown), IID
CoCreateInstance CLSID, 0, CLSCTX_INPROC_SERVER, IID, CreateIACList2
End Function

Public Function CreateIACLHistory() As IUnknown
Dim CLSID As Guid, IID As Guid
On Error Resume Next
CLSIDFromString StrPtr(CLSID_ACLHistory), CLSID
CLSIDFromString StrPtr(IID_IUnknown), IID
CoCreateInstance CLSID, 0, CLSCTX_INPROC_SERVER, IID, CreateIACLHistory
End Function

Public Function CreateIACListISF() As IUnknown
Dim CLSID As Guid, IID As Guid
On Error Resume Next
CLSIDFromString StrPtr(CLSID_ACListISF), CLSID
CLSIDFromString StrPtr(IID_IUnknown), IID
CoCreateInstance CLSID, 0, CLSCTX_INPROC_SERVER, IID, CreateIACListISF
End Function

Public Function CreateIACLMRU() As IUnknown
Dim CLSID As Guid, IID As Guid
On Error Resume Next
CLSIDFromString StrPtr(CLSID_ACLMRU), CLSID
CLSIDFromString StrPtr(IID_IUnknown), IID
CoCreateInstance CLSID, 0, CLSCTX_INPROC_SERVER, IID, CreateIACLMRU
End Function

Public Function CreateIObjMgr() As IUnknown
Dim CLSID As Guid, IID As Guid
On Error Resume Next
CLSIDFromString StrPtr(CLSID_ACLMulti), CLSID
CLSIDFromString StrPtr(IID_IObjMgr), IID
CoCreateInstance CLSID, 0, CLSCTX_INPROC_SERVER, IID, CreateIObjMgr
End Function
