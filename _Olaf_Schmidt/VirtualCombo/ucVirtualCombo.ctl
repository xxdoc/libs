VERSION 5.00
Begin VB.UserControl ucVirtualCombo 
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   KeyPreview      =   -1  'True
   ScaleHeight     =   1260
   ScaleWidth      =   3585
End
Attribute VB_Name = "ucVirtualCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Click()
Event RollUp()
Event DropDown()
Event ListMultiClick()
Event OwnerDraw(ByVal Index As Long, ByVal IsSelected As Boolean, ByVal IsComboItem As Boolean, Canvas As PictureBox, ByVal dx As Long, ByVal dy As Long)

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hwndItem As Long
  HDc As Long
  rcItem As RECT
  ItemData As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOutW& Lib "gdi32" (ByVal HDc&, ByVal x&, ByVal y&, ByVal lpString&, ByVal nCount&)
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
 
Private WithEvents oCB As VB.ComboBox, oBB As VB.PictureBox
Attribute oCB.VB_VarHelpID = -1
Private WithEvents oSL As cSubClass, WithEvents oSC As cSubClass
Attribute oSL.VB_VarHelpID = -1
Attribute oSC.VB_VarHelpID = -1
Private mItemHeight As Long, mMinVisible As Long, hWndLB As Long
Public MultiSelect As Boolean

'***** UserControl-EventHandlers
Private Sub UserControl_Initialize()
  Set oBB = Controls.Add("VB.PictureBox", "oBB")
      oBB.AutoRedraw = True: oBB.BorderStyle = 0: oBB.ScaleMode = vbPixels: oBB.BackColor = &H80000005
  hCreateHook = SetWindowsHookEx(WH_CBT, AddressOf modCreateHook.CreateHookProcForVBCombo, 0, App.ThreadID)
    Set oCB = Controls.Add("VB.ComboBox", "oCB")
  UnhookWindowsHookEx hCreateHook
  hWndLB = hComboLBox
  oCB.Visible = True
  ItemHeight = 19
  MinVisibleItems = 10
End Sub
Private Sub UserControl_Show()
  If Not oSC Is Nothing Then oSC.UnHook
  If Not oSL Is Nothing Then oSL.UnHook
  If Not Ambient.UserMode Then Exit Sub
  Set oSC = New cSubClass: oSC.Hook UserControl.hWnd
  Set oSL = New cSubClass: oSL.Hook hWndLB
End Sub
Private Sub UserControl_Hide()
  If Not oSC Is Nothing Then oSC.UnHook
  If Not oSL Is Nothing Then oSL.UnHook
End Sub
Private Sub UserControl_Resize()
  On Error Resume Next
  oCB.Move 0, 0, UserControl.Width
  UserControl.Height = oCB.Height
End Sub
Private Sub UserControl_EnterFocus()
  oCB.SetFocus
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
  If DroppedDownState And KeyAscii = vbKeySpace Then
    RaiseEvent ListMultiClick
    RedrawWindow hWndLB, 0, 0, &H101&
  End If
End Sub

'***** VB.Combo-EventHandlers
Private Sub oCB_Click()
  RaiseEvent Click
  Refresh
End Sub
Private Sub oCB_DropDown()
  RaiseEvent DropDown
End Sub

'***** SubClassing-EventHandlers
Private Sub oSL_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, Ret As Long, DefCall As Boolean)
  On Error Resume Next
    Const WM_CAPTURECHANGED = &H215, WM_LBUTTONDOWN = &H201, WM_LBUTTONUP = &H202, WM_LBUTTONDBLCLK = &H203, LB_GETCURSEL = &H188
    Static x, y, Rct As RECT, Closing As Boolean
    If Msg = WM_LBUTTONDOWN Or Msg = WM_LBUTTONUP Or Msg = WM_LBUTTONDBLCLK Then
      x = lParam And &HFFFF&: y = lParam \ &H10000
      GetClientRect hWndLB, Rct
      If MultiSelect And x > 0 And x < Rct.Right And y > 0 And y < Rct.Bottom Then
        DefCall = False
        If Msg = WM_LBUTTONUP Then RaiseEvent ListMultiClick
        RedrawWindow hWndLB, 0, 0, &H101&
      End If
    End If
    If MultiSelect And Not DroppedDownState And Msg = LB_GETCURSEL Then DefCall = False: Ret = -1
    If Msg = WM_CAPTURECHANGED Then Closing = True
    If Closing And Not DroppedDownState Then
       Closing = False: RaiseEvent RollUp
       Me.Refresh
    End If
  If Err Then Err.Clear
End Sub
Private Sub oSC_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, Ret As Long, DefCall As Boolean)
  On Error Resume Next
    Dim Drw As DRAWITEMSTRUCT
    Const WM_DRAWITEM = &H2B
    
    If Msg = WM_DRAWITEM Then
       CopyMemory Drw, ByVal lParam, Len(Drw)
       If oBB.ScaleWidth < Drw.rcItem.Right - Drw.rcItem.Left Or oBB.ScaleHeight < Drw.rcItem.Bottom - Drw.rcItem.Top Then
          oBB.Move 0, 0, ScaleX(Drw.rcItem.Right - Drw.rcItem.Left, vbPixels, ScaleMode), ScaleY(Drw.rcItem.Bottom - Drw.rcItem.Top, vbPixels, ScaleMode)
       End If
       oBB.Cls
       RaiseEvent OwnerDraw(Drw.itemID, CBool(Drw.itemState And 1), Drw.rcItem.Left > 0, oBB, Drw.rcItem.Right - Drw.rcItem.Left, Drw.rcItem.Bottom - Drw.rcItem.Top)
       BitBlt Drw.HDc, Drw.rcItem.Left, Drw.rcItem.Top, Drw.rcItem.Right - Drw.rcItem.Left, Drw.rcItem.Bottom - Drw.rcItem.Top, oBB.HDc, 0, 0, vbSrcCopy
       DefCall = False: Ret = 1
    End If
  If Err Then Err.Clear
End Sub
 
'***** Public Interface-Methods
Public Property Get ItemHeight() As Long
  ItemHeight = mItemHeight
End Property
Public Property Let ItemHeight(ByVal RHS As Long)
  mItemHeight = IIf(RHS < 4, 4, RHS)
  SendMessage oCB.hWnd, &H153, -1, ByVal mItemHeight
  SendMessage oCB.hWnd, &H153, 0, ByVal mItemHeight
  UserControl_Resize
End Property

Public Property Get MinVisibleItems() As Long
  MinVisibleItems = mMinVisible
End Property
Public Property Let MinVisibleItems(ByVal RHS As Long)
  Const CB_SETMINVISIBLE = &H1701
  mMinVisible = IIf(RHS < 4, 4, RHS)
  SendMessage oCB.hWnd, CB_SETMINVISIBLE, mMinVisible, ByVal 0&
  Dim Rct As RECT: GetClientRect oCB.hWnd, Rct
  MoveWindow oCB.hWnd, 0, 0, Rct.Right, mItemHeight * (mMinVisible + 1) + 8, 0
End Property

Public Property Get ListCount() As Long
  ListCount = oCB.ListCount
End Property
Public Property Let ListCount(ByVal RHS As Long)
  Const CB_ADDSTRING As Long = &H143, CB_RESETCONTENT As Long = &H14B
  If RHS <= 0 Then SendMessage oCB.hWnd, CB_RESETCONTENT, 0, 0&: Exit Property
  Do While RHS: RHS = RHS - 1: SendMessage oCB.hWnd, CB_ADDSTRING, 0, ByVal "": Loop
End Property

Public Property Get DroppedDownState() As Boolean
  DroppedDownState = IsWindowVisible(hWndLB) <> 0
End Property

Public Property Get ListIndex() As Long
  ListIndex = oCB.ListIndex
End Property
Public Property Let ListIndex(ByVal RHS As Long)
  oCB.ListIndex = RHS
End Property

Public Sub TextOut(x, y, ByVal S As String)
  TextOutW oBB.HDc, x, y, StrPtr(S), Len(S)
End Sub
Public Sub Refresh()
  RedrawWindow oCB.hWnd, 0, 0, &H101&
End Sub
