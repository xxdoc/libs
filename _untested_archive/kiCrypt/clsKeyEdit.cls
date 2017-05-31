VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cKeyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cKeyEdit  (clsKeyEdit.cls)
'
' Description:   These are the common edit routines you will find in most
'                word processors.  (Copy, Cut, Paste, Delete)
'
' AddIn tools    Callers Add-in v3.6 dtd 04-Sep-2016 by RD Edwards (RDE)
' for VB6:       Fantastic VB6 add-in to indentify if a routine calls
'                another routine or is called by other routines within
'                a project.  A must have tool for any VB6 programmer.
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
' 02-JUL-1998  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 11-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated ProcessNumericOnly() routine flower box with an
'              additional code suggestion.
' 18-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added NoCopyText() routine.
' 02-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed minor bugs.  See ProcessNumericOnly() and
'              ProcessAlphaNumeric() routines.
' 19-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Updated Edit_Paste() routine.  Now will remove any
'              highlighted text before pasting new data.
' ***************************************************************************
Option Explicit


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       TextBoxKeyDown
'
' Description:   Processes a KeyDown() event for a textbox.  A combination
'                of keys that have been depressed in a TextBox
'
' Syntax:        Private Sub txtLength_KeyDown(KeyCode As Integer, Shift As Integer)
'                    ' Process any key combinations
'                    mobjKeyEdit.TextBoxKeyDown txtLength, KeyCode, Shift
'                End Sub
'
' Parameters:    ctlTextBox - TextBox control
'                intKeyCode - Possible special key being pressed (Ctl, Alt, etc)
'                intShift   - Possible shift key being pressed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub TextBoxKeyDown(ByRef ctlTextBox As TextBox, _
                          ByRef intKeyCode As Integer, _
                          ByRef intShift As Integer)

    Dim CtrlDown   As Integer
    Dim PressedKey As Integer

    ' Initialize  variables
    CtrlDown = (intShift And vbCtrlMask) > 0      ' Define control key
    PressedKey = Asc(UCase$(Chr$(intKeyCode)))    ' Convert to uppercase

    ' Check to see if it is okay to make changes
    If CtrlDown And PressedKey = vbKeyX Then      ' Ctrl + X was pressed
        Edit_Cut ctlTextBox

    ElseIf CtrlDown And PressedKey = vbKeyA Then  ' Ctrl + A was pressed
        TextBoxFocus ctlTextBox

    ElseIf CtrlDown And PressedKey = vbKeyC Then  ' Ctrl + C was pressed
        Edit_Copy ctlTextBox

    ElseIf CtrlDown And PressedKey = vbKeyV Then  ' Ctrl + V was pressed
        Edit_Paste ctlTextBox

    ElseIf PressedKey = vbKeyDelete Then          ' Delete key was pressed
        Edit_Delete ctlTextBox
    End If

End Sub

' ***************************************************************************
' Routine:       TextBoxFocus
'
' Description:   Processes a GotFocus() event for a textbox.  Whenever the
'                textbox has been first entered then all the text within
'                will be highlighted.
'
' Syntax:        Private Sub txtLength_GotFocus()
'                    ' Highlight all the text in the box
'                    mobjKeyEdit.TextBoxFocus txtLength
'                End Sub
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub TextBoxFocus(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()

    ' Highlight all the text in the box
    With ctlTextBox
         .SelStart = 0             ' Start with first character in TextBox
         .SelLength = Len(.Text)   ' To last character in TextBox
    End With

End Sub

' ***************************************************************************
' Routine:       NoCopyText
'
' Description:   Prevents text from being copied from a textbox.
'
' Syntax:        Private Sub txtLength_KeyDown(KeyCode As Integer, Shift As Integer)
'                    ' Protect data from being copied
'                    mobjKeyEdit.NoCopyText txtLength, KeyCode, Shift
'                End Sub
'
' Parameters:    ctlTextBox - TextBox control
'                intKeyCode - Possible special key being pressed (Ctl, Alt, etc)
'                intShift   - Possible shift key being pressed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub NoCopyText(ByRef ctlTextBox As TextBox, _
                      ByRef intKeyCode As Integer, _
                      ByRef intShift As Integer)

    Dim CtrlDown   As Integer
    Dim PressedKey As Integer

    ' Define control key
    CtrlDown = (intShift And vbCtrlMask) > 0

    ' Convert keycode to uppercase then to decimal
    PressedKey = Asc(UCase$(Chr$(intKeyCode)))

    With ctlTextBox
        If (CtrlDown And PressedKey = vbKeyX) Or _
           (CtrlDown And PressedKey = vbKeyA) Or _
           (CtrlDown And PressedKey = vbKeyC) Or _
           (CtrlDown And PressedKey = vbKeyV) Then

            .SelStart = 0   ' Keep mouse pointer at beginning of data
            .SelLength = 0
            Clipboard.SetText vbNullString  ' Capture an empty string
            Clipboard.Clear                 ' Empty clipboard
         End If
    End With

End Sub

' ***************************************************************************
' Routine:       ProcessNumericOnly
'
' Description:   Processes a KeyPress() event for a textbox.  Numeric only
'                data.  ASCII 9 (TAB), 13 (ENTER), 8 (BACKSPACE), 48-57 (0-9)
'
'
' Syntax:        Private Sub txtLength_KeyPress(KeyAscii As Integer)
'                    ' Save only numeric and the backspace character
'                    mobjKeyEdit.ProcessNumericOnly KeyAscii
'                End Sub
'
' Parameters:    intKeyASCII - key being pressed
'
' NOTE:          To prevent someone from pasting a non-numeric value
'                into this textbox, insert the code below into the
'                textbox_Change() event. Change "txtbox" to the name
'                of the textbox control.
'
'                ' Prevent user from pasting a non-numeric value
'                ' into this textbox
'                If Not IsNumeric(txtbox.Text) Then
'                    txtbox.Text = vbNullString
'                End If
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 11-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added above note to prevent a user from pasting non-numeric
'              data into a textbox.  Got code from Masino Sinaga (04-Aug-2003)
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=47381&lngWId=1
' 02-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  ENTER key improperly recognized.  Now sends a
'              TAB key response.
' ***************************************************************************
Public Sub ProcessNumericOnly(ByRef intKeyASCII As Integer)

        Select Case intKeyASCII
               Case 9     ' Tab key
                    intKeyASCII = 0
                    SendKeys "{TAB}"
               Case 13    ' Enter key (no bell sound)
                    intKeyASCII = 0
               Case 8, 48 To 57, 127   ' Backspace, numeric keys and delete key
                    ' Good data
               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select

End Sub

' ***************************************************************************
' Routine:       ProcessAlphaNumeric
'
' Description:   Processes a KeyPress() event for a textbox.  All keyboard
'                characters.  ASCII 32-126
'
' Syntax:        Private Sub txtLength_KeyPress(KeyAscii As Integer)
'                    ' Save only alphanumeric characters
'                    mobjKeyEdit.ProcessAlphaNumeric KeyAscii
'                End Sub
'
' Parameters:    intKeyASCII - key being pressed
'                blnAcceptSpaces - [Optional]
'                     TRUE=Accept blanks spaces as input (DEFAULT)
'                     FALSE=Do not aceept any blank spaces as input
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 02-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  ENTER key improperly recognized.  Now enters a
'              Carriage Return and Linefeed.
' ***************************************************************************
Public Sub ProcessAlphaNumeric(ByRef intKeyASCII As Integer, _
                      Optional ByVal blnAcceptSpaces As Boolean = True)

    If blnAcceptSpaces Then
        Select Case intKeyASCII
               Case 9, 13   ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"

               Case 8, 32 To 127   ' Backspace and other valid data keys

               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
    Else
        ' Do not accept blank spaces
        Select Case intKeyASCII
               Case 9, 13   ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"

               Case 8, 33 To 127   ' Backspace and other valid data keys

               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
    End If

End Sub

' ***************************************************************************
' Routine:       ProcessAlphabetic
'
' Description:   Processes a KeyPress() event for a textbox.  Alphabetic
'                only data.  (A-Z, a-z)
'
' Syntax:        Private Sub txtLength_KeyPress(KeyAscii As Integer)
'                    ' Save only alphabetic only characters
'                    mobjKeyEdit.ProcessAlphabetic KeyAscii
'                End Sub
'
' Parameters:    intKeyASCII - key being pressed
'                blnUppercaseOnly - [Optional] - Allow only uppercase
'                    values to be entered.
'                    TRUE=Convert all letters to uppercase
'                    FALSE=Upper and lower case allowed (DEFAULT)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub ProcessAlphabetic(ByRef intKeyASCII As Integer, _
                    Optional ByVal blnUppercaseOnly As Boolean = False)

    If blnUppercaseOnly Then
        Select Case intKeyASCII
               Case 9, 13   ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"

               Case 8, 65 To 90, 127   ' Backspace uppercase alphabetic
                                       ' keys and delete key

               Case 97 To 122   ' Convert lowercase letters to uppercase
                    intKeyASCII = intKeyASCII - 32

               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
    Else
        ' Case does not matter
        Select Case intKeyASCII
               Case 9, 13   ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"

               Case 8, 65 To 90, 97 To 122, 127
                    ' Backspace, alphabetic keys and delete key

               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
    End If

End Sub


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Edit_Copy
'
' Description:   Copy highlighted text to the clipboard. See Keydown event
'                for the text boxes to see an example of the code calling
'                this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+C to perform a copy function.  The highlighted text
'                is then loaded into the clipboard.
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Edit_Copy(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()

    Clipboard.Clear                       ' Clear clipboard contents
    Clipboard.SetText ctlTextBox.SelText  ' Load clipboard with highlighted text

End Sub

' ***************************************************************************
' Routine:       Edit_Cut
'
' Description:   Copy highlighted text to the clipboard and then remove it
'                from the text box. See Keydown event for the text boxes to
'                see an example of the code calling this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+X to perform a cutting function.  The highlighted text
'                is then moved to the clipboard.
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Edit_Cut(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()

    Clipboard.Clear                        ' Clear clipboard contents
    Clipboard.SetText ctlTextBox.SelText   ' Load clipboard with highlighted text
    ctlTextBox.SelText = vbNullString      ' Empty TextBox contents

End Sub

' ***************************************************************************
' Routine:       Edit_Delete
'
' Description:   Copy highlighted text to the clipboard and then remove it
'                from the text box. See Keydown event for the text boxes to
'                see an example of the code calling this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+X to perform a cutting function.  The highlighted text
'                is then moved to the clipboard and the clipboard is emptied
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Edit_Delete(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()

    ctlTextBox.SelText = vbNullString   ' Remove highlighted text from TextBox

End Sub

' ***************************************************************************
' Routine:       Edit_Paste
'
' Description:   Copy whatever text is being held in the clipboard and then
'                paste it in the text box. See Keydown event for the text
'                boxes to see an example of the code calling this routine.
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Remove any highlighted text before pasting new data
' ***************************************************************************
Private Sub Edit_Paste(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()

    ' Remove highlighted text from TextBox
    If Len(ctlTextBox.SelText) > 0 Then
        ctlTextBox.SelText = vbNullString
    End If

    ctlTextBox.SelText = Clipboard.GetText()   ' Unload clipboard into TextBox

End Sub

