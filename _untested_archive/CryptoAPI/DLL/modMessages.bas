Attribute VB_Name = "modMessages"
' ***************************************************************************
'  Module:       modMessages  (modMessages.bas)
'
'  Purpose:      This module contains routines designed to provide standard
'                formatting for message boxes.  One routine can change the
'                captions on a message box.
'
' AddIn tools    Callers Add-in v3.6 dtd 04-Sep-2016 by RD Edwards (RDE)
' for VB6:       Fantastic VB6 add-in to indentify if a routine calls another
'                routine or is called by other routines within a project. A must
'                have tool for any VB6 programmer.
'                http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=74734&lngWId=1
'
'                NOTE:  Under Windows 10, if you have problems recognizing
'                a VB6 addin, try recompiling it directly into the System32
'                folder.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added custom message box routine
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added custom message box routine
' 23-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated MessageBoxH() routine on the way button captions
'                are determined.
'              - Renamed MsgBoxHookProc() to MsgboxCallBack() for easier
'                maintenance.
' 28-Aug-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic and documentation in all routines
' 03-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added option to ResponseMsg(), InfoMsg() routines to display
'              the message as a timed message. One that closes automatically
'              after a specified number of seconds.
' 25-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation and some minor tweaks.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Global constants
' ***************************************************************************
  Public Const IDOK              As Long = 1&
  Public Const IDCANCEL          As Long = 2&
  Public Const IDABORT           As Long = 3&
  Public Const IDRETRY           As Long = 4&
  Public Const IDIGNORE          As Long = 5&
  Public Const IDYES             As Long = 6&
  Public Const IDNO              As Long = 7&
  Public Const DUMMY_NUMBER      As Long = vbObjectError + 513

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const GWL_HINSTANCE    As Long = &HFFFA    ' (-6)
  Private Const HCBT_ACTIVATE    As Long = 5&
  Private Const IDPROMPT         As Long = &HFFFF&
  Private Const MB_OK            As Long = &H0&      ' one button
  Private Const MB_YESNO         As Long = &H4&      ' two buttons
  Private Const MB_YESNOCANCEL   As Long = &H3&      ' three buttons
  Private Const MB_SETFOREGROUND As Long = &H10000
  Private Const MB_TIMEDOUT      As Long = &H7D00&   ' 32000
  Private Const WH_CBT           As Long = 5&

' ***************************************************************************
' Type structures
' ***************************************************************************
  ' UDT for passing data through the hook
  Private Type MSGBOX_HOOK_PARAMS
      hwndOwner As Long
      hHook     As Long
  End Type

' ***************************************************************************
' Global Enumerations
' ***************************************************************************
  Public Enum enumCIPHER_ACTION
      eMSG_ENCRYPT   ' 0
      eMSG_DECRYPT   ' 1
  End Enum

  Public Enum enumMSGBOX_ICON
      eMSG_NOICON = 0&             ' No icon
      eMSG_ICONSTOP = 16&          ' Stop sign icon (Critical)
      eMSG_ICONQUESTION = 32&      ' Question mark icon
      eMSG_ICONEXCLAMATION = 48&   ' Exclamation mark icon
      eMSG_ICONINFORMATION = 64&   ' Information icon
  End Enum

' ***************************************************************************
' Global API Declarations
' ***************************************************************************
  ' The GetDesktopWindow function returns a handle to the desktop window.
  ' The desktop window covers the entire screen. The desktop window is
  ' the area on top of which other windows are painted.
  Public Declare Function GetDesktopWindow Lib "user32" () As Long

' ***************************************************************************
' Module API Declarations
' ***************************************************************************
  ' GetCurrentThreadId function retrieves the thread identifier of the calling
  ' thread.  The return value is the thread identifier of the calling thread.
  Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

  ' GetWindowLongPtr function retrieves information about the specified window.
  ' The function also retrieves the value at a specified offset into the extra
  ' window memory.  To write code that is compatible with both 32-bit and
  ' 64-bit versions of Windows, use GetWindowLongPtr with an alias for 32-bit
  ' version. When compiling for 32-bit Windows, GetWindowLongPtr is defined as
  ' a call to the GetWindowLong function. If the function fails, the return
  ' value is zero.
  Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" _
          (ByVal hWnd As Long, ByVal nIndex As Long) As Long

  ' Displays a modal dialog box that contains a system icon, a set of
  ' buttons, and a brief application-specific message, such as status
  ' or error information. The message box returns an integer value that
  ' indicates which button the user clicked.
  Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" _
          (ByVal hWnd As Long, ByVal lpText As String, _
          ByVal lpCaption As String, ByVal wType As Long) As Long

  ' MessageBoxTimeout function is an undocumented API which will automatically
  ' close an open msgbox after a specified amount of time.
  '
  ' You may be thinking what reassurances exist that someday Microsoft may
  ' remove this function. In all honesty, there are none; however, it is
  ' interesting to note that internally, all the documented MessageBox*
  ' functions call the MessageBoxTimeout API and simply pass 0xFFFFFFFF (-1)
  ' as the timeout period (a very long time), so the probability of it being
  ' removed is minimal.
  '
  ' https://www.codeproject.com/kb/cpp/messageboxtimeout.aspx
  Private Declare Function MessageBoxTimeout Lib "user32.dll" _
          Alias "MessageBoxTimeoutA" (ByVal hWnd As Long, _
          ByVal lpText As String, ByVal lpCaption As String, _
          ByVal uType As Long, ByVal wLanguageId As Long, _
          ByVal dwMilliseconds As Long) As Long

  ' SetDlgItemText function sets the title or text of a control in a dialog box.
  Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
          (ByVal hDlg As Long, ByVal nIDDlgItem As Long, _
          ByVal lpString As String) As Long

  ' SetWindowText function changes the text of the specified window's title
  ' bar (if it has one). If the specified window is a control, the text of
  ' the control is changed.
  Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
          (ByVal hWnd As Long, ByVal lpString As String) As Long

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
          (ByVal hHook As Long) As Long

' ***************************************************************************
' Global Variables
'                    +------------------- Global level designator
'                    |  +---------------- Data type (Boolean)
'                    |  |       |-------- Variable subname
'                    - --- --------------
' Naming standard:   g bln StopProcessing
' Variable name:     gblnStopProcessing
' ***************************************************************************
  Public gblnStopProcessing As Boolean

' ***************************************************************************
' Module Variables
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (String)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a str Captions
' Variable name:     mastrCaptions
' ***************************************************************************
  Private mlngButtonCount As Long
  Private mstrPrompt      As String
  Private mstrNewCaption  As String
  Private mastrCaptions() As String
  Private mtypMsgHook     As MSGBOX_HOOK_PARAMS


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
'  Routine:     InfoMsg
'
'  Description: Displays a Windows MsgBox with no return values.  It is
'               designed to be used where no response from the user is
'               expected other than "OK".
'
'  Parameters:  strMsg     - Message text
'               lngButtons - Numeric data designating msgbox display of
'                            buttons and/or icon
'                            Default = Exclamation mark icon with No button
'               strCaption - Optional - Msgbox caption
'                            Default = empty string
'               lngSeconds - Optional - If this is a timed message box then
'                            number of seconds to display a message is passed
'                            here.  Default = 0
'
' Example:                  +----------------------------------------------- Msgbox text message
'                           |                 |----------------------------- Msgbox button/icon codes (Optional)
'                           |       +-------------------+       +----------- Msgbox caption (Optional)
'                           |       |                   |       |       |--- Number of whole seconds to wait (Optional)
'                InfoMsg "Hello", vbOKOnly or vbInformation, "Caption", 4
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 28-Aug-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic and documentation
' 03-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to display this msgbox as a timed message
' ***************************************************************************
Public Sub InfoMsg(ByVal strMsg As String, _
          Optional ByVal lngButtons As Long = vbInformation Or vbOKOnly, _
          Optional ByVal strCaption As String = vbNullString, _
          Optional ByVal lngSeconds As Long = 0&)

    MsgboxCaption strCaption   ' Format messagebox caption

    ' Display MsgBox
    '
    '          Seconds
    ' ----------------
    ' 1 Min  = 60
    ' 1 Hour = 3600
    ' 1 Day  = 86400
    Select Case lngSeconds
           Case 1& To 600&   ' One second to ten minutes
                ' If msgbox times out, box will be closed
                TimedMsgbox strMsg, lngButtons, lngSeconds

           Case Else
                ' Will wait indefinitely for user response
                MsgBox strMsg, lngButtons, mstrNewCaption
    End Select

End Sub

' ***************************************************************************
'  Routine:     ResponseMsg
'
'  Description: Displays a standard Windows MsgBox and returns a specific
'               response code.  It is designed for when a user must make
'               a decision as to which process to perform.
'
'  Parameters:  strMsg     - Message text
'               lngButtons - Numeric data designating msgbox display of
'                            buttons and/or icon
'                            Default = Question mark icon with Yes and No
'                            buttons
'               strCaption - Optional - Msgbox caption
'                            Default = empty string
'               lngSeconds - Optional - If this is a timed message box then
'                            number of seconds to display a message is passed
'                            here.  Default = 0
'               lngDefaultButton - Optional - Default button code value. If
'                            invalid value entered then zero will be returned.
'                            Default = 1   (OK button)
'
'  Returns:     User's response to messagebox
'
'  Example:                         +-------------------------------------------------------------------------- Msgbox prompt message
'                                   |                           |---------------------------------------------- Msgbox button/icon codes (Optional)
'                                   |       +-----------------------------------------+       +---------------- Msgbox caption (Optional)
'                                   |       |                                         |       |       +-------- Number of whole seconds to wait (Optional)
'                                   |       |                                         |       |       |    |--- Default value returned if timed out (Optional)
'          lngResp = ResponseMsg("Hello", vbYesNoCancel Or vbQuestion Or vbDefaultButton2, "Caption", 3, vbNo)
'
'          Select Case lngResp
'                 Case vbYes:  ' Do something
'                 Case vbNo:   ' Do something else
'                 Case Else:   ' Exit routine
'          End Select
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 28-Aug-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic and documentation
' 03-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to display this msgbox as a timed message
' ***************************************************************************
Public Function ResponseMsg(ByVal strMsg As String, _
                   Optional ByVal lngButtons As Long = vbQuestion Or vbYesNo, _
                   Optional ByVal strCaption As String = vbNullString, _
                   Optional ByVal lngSeconds As Long = 0&, _
                   Optional ByVal lngDefaultButton As Long = vbOK) As VbMsgBoxResult

    MsgboxCaption strCaption   ' Format messagebox caption

    ' Display MsgBox
    '
    '          Seconds
    ' ----------------
    ' 1 Min  = 60
    ' 1 Hour = 3600
    ' 1 Day  = 86400
    Select Case lngSeconds
           Case 1& To 600&   ' One second to ten minutes
                ' If msgbox times out, default button code will be returned
                ResponseMsg = TimedMsgbox(strMsg, lngButtons, lngSeconds, lngDefaultButton)

           Case Else
                ' Will wait indefinitely for user response
                ResponseMsg = MsgBox(strMsg, lngButtons, mstrNewCaption)
    End Select

End Function

' ***************************************************************************
'  Routine:     ErrorMsg
'
'  Description: Displays a standard VB MsgBox formatted to display severe
'               (Usually application-type) error messages.
'
'  Parameters:  strModule  - Module where error originated
'               strRoutine - Routine where error originated
'               strMsg     - Msgbox text
'               strCaption - Optional - Msgbox caption
'                            Default = empty string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 28-Aug-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic and documentation
' ***************************************************************************
Public Sub ErrorMsg(ByVal strModule As String, _
                    ByVal strRoutine As String, _
                    ByVal strMsg As String, _
           Optional ByVal strCaption As String = vbNullString)

    Dim strFullMsg As String   ' Formatted message

    ' Verify strModule is populated
    If Len(TrimStr(strModule)) = 0 Then
       strModule = "Unknown"
    End If

    ' Verify strRoutine is populated
    If Len(TrimStr(strRoutine)) = 0 Then
       strRoutine = "Unknown"
    End If

    ' Verify strMsg is populated
    If Len(TrimStr(strMsg)) = 0 Then
       strMsg = "Unknown"
    End If

    MsgboxCaption strCaption, True   ' Format messagebox caption

    ' Format message
    strFullMsg = "Module: " & vbTab & strModule & vbCr & _
                 "Routine:" & vbTab & strRoutine & vbCr & _
                 "Error:  " & vbTab & strMsg

    ' Display MsgBox
    MsgBox strFullMsg, vbCritical Or vbOKOnly, mstrNewCaption

End Sub

' ***************************************************************************
' Routine:       MessageBoxH
'
' Description:   Displays a standard msgbox with customized captions on
'                the buttons.  Wrapper function for the MessageBox API.
'
' Reference:     VBNet - API calls for Visual Basic 6.0
'                http://vbnet.mvps.org/
'
' Parameters:    hwndForm        - Long integer system ID designating form
'                hwndWindow      - Long integer system ID designating
'                                  desktop window (API GetDesktopWindow)
'                strMsg          - Main body of text for msgbox
'                strCaption      - Caption of msgbox
'                astrBtnLabels() - String array designating button text
'                                  for up to three buttons
'                lngIcon         - Optional - Designates type of icon to use
'                                  Default - no icon
'
' Example:       ' Prepare message box display somewhere in your application.
'                '
'                ' These are the button captions,
'                ' in order, from left to right.
'                ReDim astrMsgBox(3)
'                astrMsgBox(0) = "Encrypt"
'                astrMsgBox(1) = "Decrypt"
'                astrMsgBox(2) = "Cancel"
'
'                ' Prompt user with message box
'                lngResp = MessageBoxH(Me.Hwnd, GetDesktopWindow(), _
'                                      "What do you want to do?  ", _
'                                      PGM_NAME, astrMsgBox(), eMSG_ICONQUESTION)
'                Select Case lngResp
'                       Case IDYES:    lngEncrypt = eMSG_ENCRYPT
'                       Case IDNO:     lngEncrypt = eMSG_DECRYPT
'                       Case IDCANCEL: Exit Sub
'                End Select
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Aug-2008  Randy Birch
'              http://vbnet.mvps.org/code/hooks/messageboxhook.htm
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 23-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated the way button captions are determined
' ***************************************************************************
Public Function MessageBoxH(ByVal hwndForm As Long, _
                            ByVal hwndWindow As Long, _
                            ByVal strMsg As String, _
                            ByVal strCaption As String, _
                            ByRef astrBtnLabels() As String, _
                   Optional ByVal lngIcon As enumMSGBOX_ICON = eMSG_NOICON) As Long

    Dim lngIndex    As Long
    Dim hInstance   As Long
    Dim hThreadId   As Long
    Dim lngButtonID As Long

    Erase mastrCaptions()                     ' Always start with empty arrays
    mlngButtonCount = UBound(astrBtnLabels)   ' Determine number of buttons needed
    mstrPrompt = strMsg                       ' Save msgbox text
    MsgboxCaption strCaption                  ' Format messagebox caption

    ' If array size has been exceeded then
    ' reset button count to max allowed
    If mlngButtonCount > 3 Then
        mlngButtonCount = 3
    End If

    ReDim mastrCaptions(mlngButtonCount)   ' Size array to number of captions

    ' Transfer captions to module array
    For lngIndex = 0 To mlngButtonCount - 1
        mastrCaptions(lngIndex) = astrBtnLabels(lngIndex)
    Next lngIndex

    Select Case mlngButtonCount
           Case 1: lngButtonID = MB_OK
           Case 2: lngButtonID = MB_YESNO
           Case 3: lngButtonID = MB_YESNOCANCEL
           Case Else
                MessageBoxH = IDCANCEL
                Exit Function
    End Select

    ' Set up the hook
    hInstance = GetWindowLongPtr(hwndForm, GWL_HINSTANCE)
    hThreadId = GetCurrentThreadId()

    ' Set up the MSGBOX_HOOK_PARAMS values
    ' by specifying a Windows hook as one
    ' of the params, we can intercept messages
    ' sent by Windows and thereby manipulate
    ' the dialog
    With mtypMsgHook
        .hwndOwner = hwndWindow
        .hHook = SetWindowsHookEx(WH_CBT, _
                                  AddressOf MsgboxCallBack, _
                                  hInstance, _
                                  hThreadId)
    End With

    ' Call MessageBox API and return the
    ' value as the result of the function
    MessageBoxH = MessageBox(hwndWindow, _
                             mstrPrompt, _
                             mstrNewCaption, _
                             lngButtonID Or lngIcon)

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function MsgboxCallBack(ByVal hInstance As Long, _
                                ByVal hThreadId As Long, _
                                ByVal lngNotUsed As Long) As Long

    ' Called by MessageBoxH()

    ' When the message box is about to be shown,
    ' the titlebar text, prompt message and button
    ' captions will be updated
    DoEvents
    If hInstance = HCBT_ACTIVATE Then

        ' In a HCBT_ACTIVATE message, hThreadId
        ' holds the handle to messagebox
        SetWindowText hThreadId, mstrNewCaption

        ' The ID's of the buttons on the message box
        ' correspond exactly to the values they return,
        ' so the same values can be used to identify
        ' specific buttons in a SetDlgItemText call.
        '
        ' Use default captions if array elements are empty
        Select Case mlngButtonCount
               Case 1
                    SetDlgItemText hThreadId, IDOK, IIf(Len(TrimStr(mastrCaptions(0))) > 0, mastrCaptions(0), "OK")
               Case 2
                    SetDlgItemText hThreadId, IDYES, IIf(Len(TrimStr(mastrCaptions(0))) > 0, mastrCaptions(0), "Yes")
                    SetDlgItemText hThreadId, IDNO, IIf(Len(TrimStr(mastrCaptions(1))) > 0, mastrCaptions(1), "No")
               Case 3
                    SetDlgItemText hThreadId, IDYES, IIf(Len(TrimStr(mastrCaptions(0))) > 0, mastrCaptions(0), "Yes")
                    SetDlgItemText hThreadId, IDNO, IIf(Len(TrimStr(mastrCaptions(1))) > 0, mastrCaptions(1), "No")
                    SetDlgItemText hThreadId, IDCANCEL, IIf(Len(TrimStr(mastrCaptions(2))) > 0, mastrCaptions(2), "Cancel")
        End Select

        ' Change dialog prompt text
        SetDlgItemText hThreadId, IDPROMPT, mstrPrompt

        ' Finished with dialog, release hook
        UnhookWindowsHookEx mtypMsgHook.hHook

    End If

    ' return False to let normal processing continue
    MsgboxCallBack = 0

End Function

' ***************************************************************************
'  Routine:     MsgboxCaption
'
'  Description: Formats caption text to use application title as default
'
'  Parameters:  strCaption - MsgBox caption
'               blnError   - Optional - Flag designating if something should
'                            be prefixed to messagebox caption.
'                            TRUE - Prefix "*Error*" to caption
'                            FALSE - Do not prefix data to caption (Default)
'
'  Returns:     Formatted msgbox caption
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub MsgboxCaption(ByVal strCaption As String, _
                 Optional ByVal blnErrPrefix As Boolean = False)

    ' Called by InfoMsg()
    '           ResponseMsg()
    '           ErrorMsg()

    mstrNewCaption = TrimStr(strCaption)   ' Remove unwanted characters

    ' Set caption to either input
    ' parm or the application name
    If Len(strCaption) = 0 Then

        ' Set caption default
        mstrNewCaption = App.EXEName & " v" & App.Major & "." & _
                         App.Minor & "." & App.Revision
    End If

    ' Add error prefix if requested
    If blnErrPrefix Then
        mstrNewCaption = "*Error* " & mstrNewCaption
    End If

End Sub

' ***************************************************************************
' Routine:       TimedMsgbox
'
' Description:   Display a timed msgbox
'
' Parameters:    strMsg     - Msgbox text
'                lngButtons - Numeric data designating msgbox display of
'                             buttons and/or icon
'                lngSeconds - Whole number of seconds to display msgbox
'                lngDefaultButton - Optional - Ignored unless called by
'                             ResponseMsg() routine with default button
'                             already identified
'
' Returns:       Optionally returns button selection code
'
' Example:                      +------------------------------------------------------------ Msgbox text message
'                               |                          |--------------------------------- Msgbox button/icon codes
'                               |       +--------------------------------------+    +-------- Number of whole seconds to wait
'                               |       |                                      |    |    +--- Value of default button (Optional)
'                ResponseMsg "Hello", vbOKCancel Or vbQuestion Or vbDefaultButton2, 4, vbCancel
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function TimedMsgbox(ByVal strMsg As String, _
                             ByVal lngButtons As VbMsgBoxStyle, _
                             ByVal lngSeconds As Long, _
                    Optional ByVal lngDefaultButton As Long = 0&) As Long

    ' Called by InfoMsg()
    '           ResponseMsg()

    Dim lngRetCode      As Long
    Dim lngMilliseconds As Long

    lngButtons = MB_SETFOREGROUND Or lngButtons   ' Verify msgbox becomes foreground window
    lngMilliseconds = (lngSeconds * 1000&)        ' Convert to milliseconds

    ' Display message Box with response
    lngRetCode = MessageBoxTimeout(0&, strMsg, mstrNewCaption, _
                                   lngButtons, 0&, lngMilliseconds)

    ' Evaluate user response
    '    vbOK     = 1
    '    vbCancel = 2
    '    vbAbort  = 3
    '    vbRetry  = 4
    '    vbIgnore = 5
    '    vbYes    = 6
    '    vbNo     = 7
    Select Case lngRetCode

           Case MB_TIMEDOUT   ' Time expired, msgbox closes automatically
                Select Case lngDefaultButton
                       Case 1& To 7&: TimedMsgbox = lngDefaultButton
                       Case Else:     TimedMsgbox = 0&
                End Select

           Case Else
                TimedMsgbox = lngRetCode   ' User clicked a button
    End Select

End Function

' ***************************************************************************
' Routine:       ForceEnumCase
'
' Description:   This routine exists only to ensure the CASE of these
'                constants are not altered while editing code, as can
'                happen with Enums.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Apr-2010  Karl E. Peterson
'              Customizing the Ride Part 2
'              https://visualstudiomagazine.com/articles/2010/04/12/customizing-the-ride-part-2.aspx
' 06-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Modified to support this module
' ***************************************************************************
#If False Then
Private Sub ForceEnumCase()
    ' Enum enumCIPHER_ACTION
    Const eMSG_ENCRYPT         As Long = 0&
    Const eMSG_DECRYPT         As Long = 1&

    ' Enum enumMSGBOX_ICON
    Const eMSG_NOICON          As Long = 0&    ' No icon
    Const eMSG_ICONSTOP        As Long = 16&   ' Stop sign icon (Critical)
    Const eMSG_ICONQUESTION    As Long = 32&   ' Question mark  icon
    Const eMSG_ICONEXCLAMATION As Long = 48&   ' Exclamation mark  icon
    Const eMSG_ICONINFORMATION As Long = 64&   ' Information icon
End Sub
#End If

