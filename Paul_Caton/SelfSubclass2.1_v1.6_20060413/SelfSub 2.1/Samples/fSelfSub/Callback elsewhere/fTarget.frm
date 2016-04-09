VERSION 5.00
Begin VB.Form fTarget 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Callback target"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
End
Attribute VB_Name = "fTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'* fTarget - this form receives the subclassing callbacks from fSample
'*
'*************************************************************************************************

Option Explicit

Private Type RECT
  Left              As Long
  Top               As Long
  Right             As Long
  Bottom            As Long
End Type

Private nTxtHeight  As Long                                                 'Height of a text line
Private rc          As RECT                                                 'Scrolling rectangle

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function ScrollWindowEx Lib "user32" (ByVal hwnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Form_Load()
  With Me
    nTxtHeight = .TextHeight("My")
  End With
  
  Move 7500, 1500
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    MsgBox "Note: even though the this form, the callback target, is about to be destroyed... there will be no ill-effects. The thunk detects that the callback object has been destroyed and won't callback", vbInformation
  End If
End Sub

Private Sub Form_Resize()
  rc.Right = Me.ScaleWidth
  rc.Bottom = Me.ScaleHeight
End Sub

Private Function FmtHex(ByVal nValue As Long) As String
  FmtHex = Right$("0000000" & Hex$(nValue), 8) & " "
End Function

Private Sub zWndProc_1(ByVal bBefore As Boolean, _
                       ByRef bHandled As Boolean, _
                       ByRef lReturn As Long, _
                       ByVal lng_hWnd As Long, _
                       ByVal uMsg As Long, _
                       ByVal wParam As Long, _
                       ByVal lParam As Long, _
                       ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
Const SW_INVALIDATE As Long = &H2
Const WM_PAINT      As Long = &HF
  Dim sWhen         As String

  If uMsg = WM_PAINT Then
    'If we try to display the paint message we'll just cause another paint message... vicious circle.
    Exit Sub
  End If

  If bBefore Then
    sWhen = "Before "
  Else
    sWhen = "After  "
  End If

  lParamUser = lParamUser + 1
  
  With Me
    ScrollWindowEx .hwnd, 0, -nTxtHeight, rc, rc, 0, ByVal 0&, SW_INVALIDATE
    UpdateWindow .hwnd
    .CurrentY = .ScaleHeight - nTxtHeight
    Print FmtHex(lParamUser) & sWhen & FmtHex(lReturn) & FmtHex(lng_hWnd) & FmtHex(uMsg) & FmtHex(wParam) & FmtHex(lParam)
  End With
End Sub
