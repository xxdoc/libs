VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHookContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IOleControlSite

Sub SetReference(oWb As SHDocVwCtl.WebBrowser)
    Dim oOleObj As IOleObject
    Dim oUnk As olelib.IUnknown
    Dim oFrame As IOleInPlaceFrame
    Dim oOC As IOleControl

    Set oOleObj = oWb
   
   ' Set the client site
   oOleObj.SetClientSite Me
   
   ' Activate the document
   oOleObj.DoVerb OLEIVERB_UIACTIVATE, tMSG, Me, 0, UserControl.hWnd, tRect

   ' Force the WB control to get the
   ' UA and download control properties
   Set oOC = oOleObj
   oOC.OnAmbientPropertyChange -5513
   oOC.OnAmbientPropertyChange -5512

End Sub


Private Function IOleControlSite_GetExtendedControl() As Object

End Function

Private Sub IOleControlSite_LockInPlaceActive(ByVal fLock As olelib.BOOL)

End Sub

Private Sub IOleControlSite_OnControlInfoChanged()

End Sub

Private Sub IOleControlSite_OnFocus(ByVal fGotFocus As olelib.BOOL)

End Sub

Private Sub IOleControlSite_ShowPropertyFrame()

End Sub

Private Sub IOleControlSite_TransformCoords(pPtlHimetric As olelib.POINT, pPtfContainer As olelib.POINTF, ByVal dwFlags As olelib.XFORMCOORDS)

End Sub

Private Sub IOleControlSite_TranslateAccelerator(pMsg As olelib.MSG, ByVal grfModifiers As olelib.KEYMODIFIERS)

End Sub
