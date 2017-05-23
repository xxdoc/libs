VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWbExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDocHostUIHandler

Event EditUrlBeforeNavigate(url As String)
Event OnContextMenu(cancel As Boolean)
Event KeyPress(keyCode As Integer, Accelerator As Integer, BlockInput As Boolean)
Event ShortCutKey(cmdId As WbCommands, BlockCommand As Boolean)
Event GetExternal(oIDispatch As Object)


Private myUnknownCommand As Long
Private Const E_NOTIMPL = -2147467263 '(&H80004001)
Private IID_CommandGroup As ieKit.UUID
Private Const sIID_CommandGroup As String = "{DE4BA900-59CA-11CF-9592-444553540000}"

Public WbAttributes As HostAttributes

Enum HostAttributes
   haDisableSelections = DOCHOSTUIFLAG_DIALOG
   haNo3DBorder = DOCHOSTUIFLAG_NO3DBORDER
   haNoScrollBars = DOCHOSTUIFLAG_SCROLL_NO
   haUseFlatScrollBars = DOCHOSTUIFLAG_FLAT_SCROLLBAR
   haInPlaceNavigation = DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION
End Enum

Enum WbCommands
    wbPrint = 27
    wbFind = 67
    wbPaste = 26
    wbSelectAll = 31
    wbGoBack = 2282
    wbGoForw = 2283
    wbCopy = 15
End Enum

Enum ContextMenuTargets
    cmDefault = 0
    cmImage = 1
    cmControl = 2
    cmTable = 3
    cmSelection = 4
    cmLink = 5
    cmAnchor = 6
    cmImgDynsrc = 7
    cmImgart = 8
End Enum
    
Public ContextMenuTargetType As ContextMenuTargets
Public ContextMenuTargetHtmlObject As Object
    
Function HookWebBrowser(oWb As Object) As Boolean
 On Error GoTo hell
   Dim iCustDoc As ICustomDoc
   Set iCustDoc = oWb.Document
   iCustDoc.SetUIHandler Me
   HookWebBrowser = True
Exit Function
hell:
End Function

Private Sub Class_Initialize()
    CLSIDFromString sIID_CommandGroup, IID_CommandGroup
End Sub

Private Sub IDocHostUIHandler_EnableModeless(ByVal fEnable As ieKit.BOOL)
    Err.Raise E_NOTIMPL
End Sub

Private Function IDocHostUIHandler_FilterDataObject(ByVal pDO As ieKit.IDataObject) As ieKit.IDataObject
    Err.Raise S_OK
End Function

Private Function IDocHostUIHandler_GetDropTarget(ByVal pDropTarget As ieKit.IDropTarget) As ieKit.IDropTarget
    Err.Raise E_NOTIMPL
End Function

Private Function IDocHostUIHandler_GetExternal() As Object
    RaiseEvent GetExternal(IDocHostUIHandler_GetExternal)
End Function

Private Sub IDocHostUIHandler_GetHostInfo(pInfo As ieKit.DOCHOSTUIINFO)
    pInfo.dwFlags = WbAttributes
End Sub

Private Sub IDocHostUIHandler_GetOptionKeyPath(pOLESTRchKey As Long, ByVal dw As Long)
    Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_HideUI()
    Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_OnDocWindowActivate(ByVal fActivate As ieKit.BOOL)
    Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_OnFrameWindowActivate(ByVal fActivate As ieKit.BOOL)
    Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_ResizeBorder(prcBorder As ieKit.RECT, ByVal pUIWindow As ieKit.IOleInPlaceUIWindow, ByVal fRameWindow As ieKit.BOOL)
    Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_ShowContextMenu(ByVal dwContext As ieKit.ContextMenuTarget, pPOINT As ieKit.POINT, ByVal pCommandTarget As ieKit.IOleCommandTarget, ByVal HTMLTagElement As Object)
    Dim disable As Boolean
    
    ContextMenuTargetType = dwContext
    Set ContextMenuTargetHtmlObject = HTMLTagElement
    
    RaiseEvent OnContextMenu(disable)
    
    ContextMenuTargetType = 0
    Set ContextMenuTargetHtmlObject = Nothing
    
    If Not disable Then Err.Raise E_NOTIMPL
    
End Sub

Private Sub IDocHostUIHandler_ShowUI(ByVal dwID As Long, ByVal pActiveObject As ieKit.IOleInPlaceActiveObject, ByVal pCommandTarget As ieKit.IOleCommandTarget, ByVal pFrame As ieKit.IOleInPlaceFrame, ByVal pDoc As ieKit.IOleInPlaceUIWindow)
    Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_TranslateAccelerator(lpmsg As ieKit.MSG, pguidCmdGroup As ieKit.UUID, ByVal nCmdID As Long)
   Dim Accelerate As Integer, Char As Integer, block As Boolean
   
   If nCmdID <> 0 Then
       If IsEqualGUID(pguidCmdGroup, IID_CommandGroup) Then
            RaiseEvent ShortCutKey(nCmdID, block)
       End If
   Else 'keyboard event
            If lpmsg.message = &H102 Then 'WMCHAR
                Char = lpmsg.wParam
                
                If GetAsyncKeyState(vbKeyShift) Then setBit Accelerate, vbShiftMask
                If GetAsyncKeyState(vbKeyMenu) Then setBit Accelerate, vbAltMask
                If GetAsyncKeyState(vbKeyControl) Then setBit Accelerate, vbCtrlMask
            
                RaiseEvent KeyPress(Char, Accelerate, block)
                lpmsg.wParam = Char
            End If
   End If

  If block = True Then Exit Sub
  Err.Raise E_NOTIMPL
   
End Sub

Private Function IDocHostUIHandler_TranslateUrl(ByVal dwTranslate As Long, ByVal pchURLIn As Long) As Long
   Dim sUrl As String
   sUrl = SysAllocString(pchURLIn) 'Get URL from pointer
   RaiseEvent EditUrlBeforeNavigate(sUrl)
   IDocHostUIHandler_TranslateUrl = Str2Ptr(sUrl)
End Function

Private Sub IDocHostUIHandler_UpdateUI()
    Err.Raise E_NOTIMPL
End Sub

Private Function Str2Ptr(ByRef Str As String) As Long
   Str2Ptr = CoTaskMemAlloc(LenB(Str) + 2) ' Allocate memory for the string
   MoveMemory ByVal Str2Ptr, ByVal StrPtr(Str), LenB(Str) + 2 ' Copy the string
End Function

Private Sub setBit(v, mask)
    v = v Or mask
End Sub
