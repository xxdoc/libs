VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   5636
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements olelib.IOleClientSite
Implements olelib2.IOleInPlaceSite
Implements olelib2.IServiceProvider
Implements olelib.IInternetSecurityManager

Implements olelib.IDocHostUIHandler

Dim m_ozm As olelib.IInternetZoneManager
Private IID_InternetSecurityManager As olelib.UUID
Private Const sIID_InternetSecurityManager = "{79eac9ee-baf9-11ce-8c82-00aa004ba90b}"

Public Sub nav(url As String)
    wb.Navigate url
End Sub
Public Property Get BrowseMode() As Boolean
   BrowseMode = True
End Property

Public Property Let BrowseMode(New_BrowseMode As Boolean)
Dim oOC As IOleControl

   'm_bBrowseMode = New_BrowseMode

   ' Get the WB IOleControl
   Set oOC = wb.Document

   ' Notify the WB control that
   ' the property was changed
   oOC.OnAmbientPropertyChange AMBIENT_DISPIDS.DISPID_AMBIENT_USERMODE

End Property

Public Property Get DownloadCtrl() As DownloadCtrlFlags

   DownloadCtrl = DLCTL_NO_SCRIPTS

End Property

Public Property Let DownloadCtrl(ByVal NewFlags As DownloadCtrlFlags)
Dim oOC As IOleControl

   'm_lDownloadCtrl = NewFlags

   ' Get the WB IOleControl
   Set oOC = wb.Document

   ' Notify the WB control that
   ' the property was changed
   oOC.OnAmbientPropertyChange -5512

End Property

Public Property Get UserAgent() As String

   UserAgent = "I am Mozilla Hear me Roar!"

End Property

Public Property Let UserAgent(ByVal New_UA As String)
Dim oOC As IOleControl
 
   'm_sUserAgent = New_UA

   ' Get the WB IOleControl
   Set oOC = wb.Parent

   ' Notify the WB control that
   ' the property was changed
   oOC.OnAmbientPropertyChange -5513

End Property


Sub HookWb()
    
    Dim oOleObj As IOleObject
    Dim oOC As IOleControl

   ' Get the IOleObject interface
   Set oOleObj = wb.Application

   ' Set the client site
   oOleObj.SetClientSite Me
   
   ' Force the WB control to get the UA and download control properties
   Set oOC = oOleObj
   oOC.OnAmbientPropertyChange -5513
   oOC.OnAmbientPropertyChange -5512
   
   CoInternetCreateZoneManager Nothing, m_ozm, 0
   
   Dim iCustDoc As ICustomDoc
   Set iCustDoc = oWb.Document
   iCustDoc.SetUIHandler Me

End Sub


Private Sub IDocHostUIHandler_EnableModeless(ByVal fEnable As olelib.BOOL)
Err.Raise E_NOTIMPL
End Sub

Private Function IDocHostUIHandler_FilterDataObject(ByVal pDO As olelib.IDataObject) As olelib.IDataObject
Err.Raise E_NOTIMPL
End Function

Private Function IDocHostUIHandler_GetDropTarget(ByVal pDropTarget As olelib.IDropTarget) As olelib.IDropTarget
Err.Raise E_NOTIMPL
End Function

Private Function IDocHostUIHandler_GetExternal() As Object
Err.Raise E_NOTIMPL
End Function

Private Sub IDocHostUIHandler_GetHostInfo(pInfo As olelib.DOCHOSTUIINFO)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_GetOptionKeyPath(pOLESTRchKey As Long, ByVal dw As Long)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_HideUI()
Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_OnDocWindowActivate(ByVal fActivate As olelib.BOOL)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_OnFrameWindowActivate(ByVal fActivate As olelib.BOOL)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_ResizeBorder(prcBorder As olelib.RECT, ByVal pUIWindow As olelib.IOleInPlaceUIWindow, ByVal fRameWindow As olelib.BOOL)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_ShowContextMenu(ByVal dwContext As olelib.ContextMenuTarget, pPOINT As olelib.POINT, ByVal pCommandTarget As olelib.IOleCommandTarget, ByVal HTMLTagElement As Object)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_ShowUI(ByVal dwID As Long, ByVal pActiveObject As olelib.IOleInPlaceActiveObject, ByVal pCommandTarget As olelib.IOleCommandTarget, ByVal pFrame As olelib.IOleInPlaceFrame, ByVal pDoc As olelib.IOleInPlaceUIWindow)
Err.Raise E_NOTIMPL
End Sub

Private Sub IDocHostUIHandler_TranslateAccelerator(lpmsg As olelib.MSG, pguidCmdGroup As olelib.UUID, ByVal nCmdID As Long)
Err.Raise E_NOTIMPL
End Sub

Private Function IDocHostUIHandler_TranslateUrl(ByVal dwTranslate As Long, ByVal pchURLIn As Long) As Long
Err.Raise E_NOTIMPL
End Function

Private Sub IDocHostUIHandler_UpdateUI()
Err.Raise E_NOTIMPL
End Sub

Private Function IOleClientSite_GetContainer() As olelib.IOleContainer
    Err.Raise E_NOTIMPL
End Function

Private Function IOleClientSite_GetMoniker(ByVal dwAssign As olelib.OLEGETMONIKER, ByVal dwWhichMoniker As olelib.OLEWHICHMK) As olelib.IMoniker
Err.Raise E_NOTIMPL
End Function

Private Sub IOleClientSite_OnShowWindow(ByVal fShow As olelib.BOOL)
Err.Raise E_NOTIMPL
End Sub

Private Sub IOleClientSite_RequestNewObjectLayout()
Err.Raise E_NOTIMPL
End Sub

Private Sub IOleClientSite_SaveObject()
 
End Sub

Private Sub IOleClientSite_ShowObject()
    Err.Raise E_NOTIMPL
End Sub

Private Sub IOleInPlaceSite_CanInPlaceActivate()

End Sub

Private Sub IOleInPlaceSite_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)

End Sub

Private Sub IOleInPlaceSite_DeactivateAndUndo()

End Sub

Private Sub IOleInPlaceSite_DiscardUndoState()

End Sub

Private Function IOleInPlaceSite_GetWindow() As Long
    IOleInPlaceSite_GetWindow = UserControl.hwnd
End Function

Private Sub IOleInPlaceSite_GetWindowContext(ppFrame As olelib.IOleInPlaceFrame, ppDoc As olelib.IOleInPlaceUIWindow, lprcPosRect As olelib.RECT, lprcClipRect As olelib.RECT, lpFrameInfo As olelib.OLEINPLACEFRAMEINFO)
    Set ppFrame = Me
   Set ppDoc = Me
   
   lpFrameInfo.hwndFrame = UserControl.hwnd
End Sub

Private Sub IOleInPlaceSite_OnInPlaceActivate()

End Sub

Private Sub IOleInPlaceSite_OnInPlaceDeactivate()

End Sub

Private Sub IOleInPlaceSite_OnPosRectChange(lprcPosRect As olelib.RECT)

End Sub

Private Sub IOleInPlaceSite_OnUIActivate()

End Sub

Private Sub IOleInPlaceSite_OnUIDeactivate(ByVal fUndoable As olelib.BOOL)

End Sub

Private Sub IOleInPlaceSite_Scroll(ByVal scrollX As Long, ByVal scrollY As Long)

End Sub

Private Sub IServiceProvider_QueryService(guidService As olelib.UUID, riid As olelib.UUID, ppvObject As Long)

   If IsEqualGUID(guidService, IID_InternetSecurityManager) Then

      Dim oISM As IInternetSecurityManager ''

      ' Increment the reference count
      'pvAddRefMe

      Set oISM = Me

      ' Return this object
      MoveMemory ppvObject, oISM, 4&

   Else
      Dim mystr As String
      
      Debug.Print StringFromGUID2(guidService, SysAllocString(mystr), 256)
      
      ' The service or interface is
      ' not supported
      Err.Raise E_NOINTERFACE

   End If

End Sub




Private Sub IInternetSecurityManager_GetSecurityId(ByVal pwszUrl As Long, ByVal pbSecurityId As Long, pcbSecurityId As Long, ByVal dwReserved As Long)
   Err.Raise INET_E_DEFAULT_ACTION
End Sub

Private Function IInternetSecurityManager_GetSecuritySite() As olelib.IInternetSecurityMgrSite
   Err.Raise INET_E_DEFAULT_ACTION
End Function

Private Sub IInternetSecurityManager_GetZoneMappings(ByVal dwZone As Long, ppenumString As olelib.IEnumString, ByVal dwFlags As Long)
   Err.Raise INET_E_DEFAULT_ACTION
End Sub

Private Sub IInternetSecurityManager_MapUrlToZone(ByVal pwszUrl As Long, pdwZone As Long, ByVal dwFlags As Long)
   Err.Raise INET_E_DEFAULT_ACTION
End Sub

Private Sub IInternetSecurityManager_ProcessUrlAction( _
      ByVal pwszUrl As Long, _
      ByVal dwAction As URLACTIONS, _
      ByVal pPolicy As Long, _
      ByVal cbPolicy As Long, _
      pContext As Byte, _
      ByVal cbContext As Long, _
      ByVal dwFlags As olelib.PUAF, _
      ByVal dwReserved As Long)

Dim lPolicy As olelib.URLPOLICIES
Dim abPolicy(0 To 3) As Byte
Dim uz As URLZONE
uz = URLZONE_INTERNET

   ' Get the policy for the control security zone
   m_ozm.GetZoneActionPolicy uz, dwAction, abPolicy(0), 4&, URLZONEREG_DEFAULT
   MoveMemory lPolicy, abPolicy(0), 4&

   ' Ask the container for a policy. This allows the container to
   ' overwrite the policies for the  selected security zone
   'RaiseEvent ProcessAction(SysAllocString(pwszUrl), dwAction, lPolicy)

   ' Copy the policy to the pointer
   MoveMemory ByVal pPolicy, lPolicy, 4&

End Sub

Private Sub IInternetSecurityManager_QueryCustomPolicy(ByVal pwszUrl As Long, guidKey As olelib.UUID, ppPolicy As Long, pcbPolicy As Long, pContext As Byte, ByVal cbContext As Long, Optional ByVal dwReserved As Long = 0&)
   Err.Raise INET_E_DEFAULT_ACTION
End Sub


Private Sub IInternetSecurityManager_SetSecuritySite(ByVal pSite As olelib.IInternetSecurityMgrSite)
   Err.Raise INET_E_DEFAULT_ACTION
End Sub

Private Sub IInternetSecurityManager_SetZoneMapping(ByVal dwZone As Long, ByVal lpszPattern As Long, ByVal dwFlags As olelib.SZM_FLAGS)
   Err.Raise INET_E_DEFAULT_ACTION
End Sub















Private Sub UserControl_Initialize()
    CLSIDFromString sIID_InternetSecurityManager, IID_InternetSecurityManager
End Sub

Private Sub pvAddRefMe()
    Dim oUnk As olelib.IUnknown

   Set oUnk = Me
   oUnk.AddRef

End Sub
