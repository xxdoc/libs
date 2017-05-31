Attribute VB_Name = "modCentering"
' ***************************************************************************
' Module:        modCentering (modCentering.bas)
'
' Description:   Centering routines which include:
'                    - Center a caption on a form
'                    - Center report text on a page
'                    - Center one form on top of another
'                    - Center a form on the screen
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 19-Feb-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added CenterCaption(), CenterReportText() routines
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME      As String = "modCentering"
  Private Const GWL_STYLE        As Long = (-16)   ' Retrieves window styles
  Private Const GWL_HINSTANCE    As Long = (-6)
  Private Const HCBT_ACTIVATE    As Long = 5
  Private Const SWP_NOSIZE       As Long = &H1&
  Private Const SWP_NOMOVE       As Long = &H2&
  Private Const SWP_NOZORDER     As Long = &H4&
  Private Const SWP_NOACTIVATE   As Long = &H10&
  Private Const SWP_FRAMECHANGED As Long = &H20&
  Private Const WH_CBT           As Long = 5

  Private Const REDRAW_FLAGS     As Long = SWP_NOSIZE Or _
                                           SWP_NOMOVE Or _
                                           SWP_NOZORDER Or _
                                           SWP_NOACTIVATE Or _
                                           SWP_FRAMECHANGED

' ***************************************************************************
' Type Structures
' ***************************************************************************
  Private Type RECT
      Left   As Long   ' Left side of form
      Top    As Long   ' Top side of form
      Right  As Long   ' Right side of form
      Bottom As Long   ' Bottom side of form
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' GetWindowRect function retrieves the dimensions of the bounding
  ' rectangle of the specified window.  The dimensions are given in
  ' screen coordinates that are relative to the upper-left corner of
  ' the screen.
  Private Declare Function GetWindowRect Lib "user32" _
          (ByVal hWnd As Long, lpRect As RECT) As Long

  ' GetWindowLongPtr function retrieves information about the specified window.
  ' The function also retrieves the value at a specified offset into the extra
  ' window memory.  To write code that is compatible with both 32-bit and
  ' 64-bit versions of Windows, use GetWindowLongPtr with an alias for 32-bit
  ' version. When compiling for 32-bit Windows, GetWindowLongPtr is defined as
  ' a call to the GetWindowLong function. If the function fails, the return
  ' value is zero.
  Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" _
          (ByVal hWnd As Long, ByVal nIndex As Long) As Long

  ' SetWindowLongPtr function changes an attribute of the specified window.
  ' The function also retrieves the value at a specified offset into the extra
  ' window memory.  To write code that is compatible with both 32-bit and
  ' 64-bit versions of Windows, use SetWindowLongPtr with an alias for 32-bit
  ' version. When compiling for 32-bit Windows, SetWindowLongPtr is defined as
  ' a call to the SetWindowLong function. If the function fails, the return
  ' value is zero.
  Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" _
          (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

  ' SetWindowPos function changes the size, position, and Z order of a child,
  ' pop-up, or top-level window. These windows are ordered according to their
  ' appearance on the screen. The topmost window receives the highest rank
  ' and is the first window in the Z order. If the function fails, the return
  ' value is zero.
  Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
          ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

  ' GetCurrentThreadId function retrieves the thread identifier of the calling
  ' thread.  The return value is the thread identifier of the calling thread.
  Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

  ' SetWindowsHookEx function installs an application-defined hook procedure
  ' into a hook chain.  If the function succeeds, the return value is the
  ' handle to the hook procedure.
  Private Declare Function SetWindowsHookEx Lib "user32" _
          Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
          ByVal hmod As Long, ByVal dwThreadId As Long) As Long

  ' UnhookWindowsHookEx function removes a hook procedure installed in a hook
  ' chain by the SetWindowsHookEx function.  If the function succeeds, the
  ' return value is nonzero.
  Private Declare Function UnhookWindowsHookEx Lib "user32" _
          (ByVal mlngHook As Long) As Long

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Long)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m lng Hook
' Variable name:     mlngHook
' ***************************************************************************
  Private mlngHook          As Long   ' Used for centering formes
  Private mlngFrmTopHwnd    As Long   ' Top form handle (To be centered)
  Private mlngFrmBottomHwnd As Long   ' Bottom form handle (To be centered on)


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CenterCaption
'
' Description:   Centers a caption on a form.
'
' Parameters:    frmForm - Name of form whose caption is to be centered
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Feb-2015  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine.
' ***************************************************************************
Public Sub CenterCaption(ByRef frmName As Form)

    Dim lngHwnd As Long

    On Error Resume Next

    With frmName

        ' Retrieve form style
        lngHwnd = GetWindowLongPtr(.hWnd, GWL_STYLE)

        ' Change form offset
        If SetWindowLongPtr(.hWnd, GWL_STYLE, lngHwnd) <> 0 Then

            ' Update form caption
            SetWindowPos .hWnd, 0&, 0&, 0&, 0&, 0&, REDRAW_FLAGS
        End If

    End With

CenterCaption_CleanUp:
    On Error GoTo 0   ' Nullify this error trap

End Sub

' ***************************************************************************
' Routine:       CenterReportText
'
' Description:   Center text on a line
'
' Parameters:    lngLineLength - Length of report line
'                strLeftSide - Optional - String of data to remain at left
'                    most end of output string.  (ex:  "25-Dec-2010")
'                    Default = vbNullString
'                strMiddle - Optional - String of data to be centered.
'                    (ex:  "My name and email")
'                    Default = vbNullString
'                strRightSide - Optional - String of data to remain at right
'                    most end of output string.  (ex:  "Page 1")
'                    Default = vbNullString
'
' Returns:       Formatted text
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 16-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added better parameter evaluation with error messages
' 11-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Made line length parameter mandatory
' ***************************************************************************
Public Function CenterReportText(ByVal lngLineLength As Long, _
                        Optional ByVal strLeftSide As String = vbNullString, _
                        Optional ByVal strMiddle As String = vbNullString, _
                        Optional ByVal strRightSide As String = vbNullString) As String

    Dim lngDataLength As Long

    Const ROUTINE_NAME As String = "CenterReportText"

    CenterReportText = vbNullString   ' Verify return string is empty

    ' If no line length, data cannot be centered
    If lngLineLength < 1 Then
        InfoMsg "Line length must be a positive number." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    ' Remove leading and trailing blank spaces
    strLeftSide = TrimStr(strLeftSide)
    strMiddle = TrimStr(strMiddle)
    strRightSide = TrimStr(strRightSide)

    ' Capture data length
    lngDataLength = Len(strLeftSide & strMiddle & strRightSide)

    Select Case lngDataLength

           Case Is < 1   ' If no data to process then leave
                InfoMsg "No data to process." & vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function

           Case Is > lngLineLength   ' If too much data then leave
                InfoMsg "Line length must be equal to or greater than data length." & _
                        vbNewLine & "Line length:  " & CStr(lngLineLength) & _
                        vbNewLine & "Data length:  " & CStr(lngDataLength) & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function
    End Select

    ' Add a blank space to beginning and end of
    ' middle string of data until line length
    ' requirement has been met or exceeded
    Do While Len(strMiddle) < lngLineLength
        strMiddle = Chr$(32) & strMiddle & Chr$(32)
    Loop

    ' Verify string length equals line length
    strMiddle = Left$(strMiddle, lngLineLength)

    ' If data is available then overlay far left side
    If Len(strLeftSide) > 0 Then
        Mid$(strMiddle, 1, Len(strLeftSide)) = strLeftSide
    End If

    ' If data is available then overlay far right side
    If Len(strRightSide) > 0 Then
        Mid$(strMiddle, (lngLineLength - Len(strRightSide)) + 1, Len(strRightSide)) = strRightSide
    End If

    ' Remove any excess trailing blanks because
    ' only leading blanks are needed to push data
    ' to middle of line.
    CenterReportText = RTrim$(strMiddle)

End Function

' ***************************************************************************
' Routine:       CenterForm
'
' Description:   Center form on screen or on top of another form using
'                a call-back function.
'
'                The easiest way to center a form on the screen is to
'                enter the following code in the Form_Load() event of
'                a form to be centered.
'
'                With frmMain
'                    .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
'                End With
'
' Parameters:    frmTop - Form to be centered (ex: frmAbout)
'                frmBottom - Optional - Name of form to be centered
'                    on. (ex: frmMain)  If missing, top form is centered
'                    on screen.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Sep-1999  Paul Mather
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
' 30-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub CenterForm(ByRef frmTop As Form, _
             Optional ByRef frmBottom As Form = Nothing)

    Dim lngThread       As Long
    Dim lngWinScreen    As Long
    Dim lngBottomHwnd   As Long
    Dim blnCenterOnForm As Boolean

    mlngHook = 0                   ' Preset to Not Found
    mlngFrmTopHwnd = frmTop.hWnd   ' Top form handle

    ' Determine if form is to be centered
    ' on screen or on top of another form
    If frmBottom Is Nothing Then
        blnCenterOnForm = False   ' Center on window screen
    Else
        blnCenterOnForm = True    ' Center on another form
    End If

    Select Case blnCenterOnForm

           Case True    ' Center form on top of another form
                mlngFrmBottomHwnd = frmBottom.hWnd   ' Bottom form handle
                lngBottomHwnd = GetWindowLongPtr(mlngFrmBottomHwnd, GWL_HINSTANCE)
                lngThread = GetCurrentThreadId()
                mlngHook = SetWindowsHookEx(WH_CBT, AddressOf CenterOnForm, lngBottomHwnd, lngThread)

           Case False   ' Center form on screen
                lngWinScreen = GetWindowLongPtr(frmTop.hWnd, GWL_HINSTANCE)
                lngThread = GetCurrentThreadId()
                mlngHook = SetWindowsHookEx(WH_CBT, AddressOf CenterOnScreen, lngWinScreen, lngThread)
    End Select

    DoEvents

End Sub


' ***************************************************************************
' ****               Internal procedures & functions                     ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CenterOnForm
'
' Description:   Center one form on top of another
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Sep-1999  Paul Mather
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
' 30-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function CenterOnForm(ByVal lngMsg As Long, _
                              ByVal lngBottomHwnd As Long, _
                              ByVal lngThread As Long) As Long

    ' Called by CenterForm()

    Dim lngPosX   As Long   ' Upper left corner position from left side of screen
    Dim lngPosY   As Long   ' Upper left corner position from top of screen
    Dim typTOP    As RECT   ' Top form coordinates
    Dim typBOTTOM As RECT   ' Bottom form coordinates

    ' Show on form centered over another form
    If lngMsg = HCBT_ACTIVATE Then

        ' Get form coordinates
        GetWindowRect mlngFrmBottomHwnd, typBOTTOM   ' Capture bottom form dimensions
        GetWindowRect mlngFrmTopHwnd, typTOP         ' Capture top form dimensions

        ' Calculate where to locate form based on
        ' calling form location and dimensions
        lngPosX = (typBOTTOM.Left + (typBOTTOM.Right - typBOTTOM.Left) / 2) - ((typTOP.Right - typTOP.Left) / 2)
        lngPosY = (typBOTTOM.Top + (typBOTTOM.Bottom - typBOTTOM.Top) / 2) - ((typTOP.Bottom - typTOP.Top) / 2)

        ' Move top form to new location
        SetWindowPos lngBottomHwnd, 0, lngPosX, lngPosY, 0, 0, _
                     SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE

        UnhookWindowsHookEx mlngHook   ' Release CBT hook

    End If

    CenterOnForm = 0

End Function

' ***************************************************************************
' Routine:       CenterOnScreen
'
' Description:   Center a form on screen regardless of monitor size
'
'                The easiest way to center a form on the screen is to
'                enter the following code in the Form_Load() event of
'                a form to be centered.
'
'                With frmMain
'                    .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
'                End With
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Sep-1999  Paul Mather
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
' 30-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function CenterOnScreen(ByVal lngMsg As Long, _
                                ByVal lngWinScreen As Long, _
                                ByVal lngThread As Long) As Long

    ' Called by CenterForm()

    Dim lngPosX As Long   ' Upper left corner position from left side of screen
    Dim lngPosY As Long   ' Upper left corner position from top of screen
    Dim typTOP  As RECT   ' Top form coordinates

    ' Show form centered on screen
    If lngMsg = HCBT_ACTIVATE Then

        GetWindowRect lngWinScreen, typTOP   ' Capture window dimensions

        ' Calculate where to locate form
        lngPosX = (Screen.Width / Screen.TwipsPerPixelX / 2) - ((typTOP.Right - typTOP.Left) / 2)
        lngPosY = (Screen.Height / Screen.TwipsPerPixelY / 2) - ((typTOP.Bottom - typTOP.Top) / 2)

        ' Move form to center of screen
        SetWindowPos lngWinScreen, 0, lngPosX, lngPosY, 0, 0, _
                     SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE

        UnhookWindowsHookEx mlngHook   ' Release CBT hook

    End If

    CenterOnScreen = 0

End Function

