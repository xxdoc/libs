Attribute VB_Name = "modProcesses"
' ***************************************************************************
' Module:        modProcesses  (modProcesses.bas)
'
' Purpose:       Find and/or Stop designated processes completely.
'                Find a specific parent window with a complete or partial
'                caption and any child windows or all parent and child
'                windows and return the information in an array.
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
' 10-Feb-2014  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote module
' 10-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Added FindHandleByExe() routine.
'              - Rewrote FindProcessByHandle() routine.
' 13-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Added GetParentChildWindows(), GetParentWindows(),
'              GetChildWindows() routines
' 25-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to capture Parent windows only.  See
'              GetParentChildWindows(), GetParentWindows() routines.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MAX_SIZE           As Long = 260
  Private Const ARRY_SIZE          As Long = 100
  Private Const WM_GETTEXT         As Long = &HD
  Private Const WM_GETTEXTLENGTH   As Long = &HE
  Private Const GW_HWNDNEXT        As Long = 2
  Private Const PROCESS_QUERY_INFO As Long = &H400
  Private Const PROCESS_VM_READ    As Long = &H10
  Private Const TH32CS_SNAPPROCESS As Long = &H2&

' ****************************************************************************
' Type Structures
' ****************************************************************************
  Private Type PROCESSENTRY32
      dwSize              As Long
      cntUsage            As Long
      th32ProcessID       As Long
      th32DefaultHeapID   As Long
      th32ModuleID        As Long
      cntThreads          As Long
      th32ParentProcessID As Long
      pcPriClassBase      As Long
      dwFlags             As Long
      szExeFile           As String * 260
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' ZeroMemory function fills a block of memory with zeros.
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

  ' Always close an objects handle if it is not being used.
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long

  ' OpenProcess function returns a handle of an existing process object.
  ' If the function succeeds, the return value is an open handle of the
  ' specified process.
  Private Declare Function OpenProcess Lib "kernel32" _
          (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
          ByVal dwProcessId As Long) As Long

  ' The TerminateProcess() function is used to unconditionally cause a
  ' process to exit and not save anything.
  Private Declare Function TerminateProcess Lib "kernel32" _
          (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

  ' CreateToolhelpSnapshot Takes a snapshot of the specified processes,
  ' as well as the heaps, modules, and threads used by these processes.
  Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
          (ByVal lFlags As Long, lProcessID As Long) As Long

  ' Process32First function retrieves information about the first process
  ' encountered in a system snapshot.
  Private Declare Function Process32First Lib "kernel32" _
          (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

  ' Process32Next function retrieves information about the next process
  ' recorded in a system snapshot.
  Private Declare Function Process32Next Lib "kernel32" _
          (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

  ' FindWindow function retrieves the handle to the top-level window
  ' whose class name and window name match the specified strings.
  ' This function does not search child windows.
  Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
          (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

  ' GetWindow function retrieves the handle of a window that has the
  ' specified relationship (Z order or owner) to the specified window.
  Private Declare Function GetWindow Lib "user32" _
          (ByVal hWnd As Long, ByVal wCmd As Long) As Long

  ' GetWindowThreadProcessId function retrieves the identifier of
  ' the thread that created the specified window and, optionally,
  ' the identifier of the process that created the window.
  Private Declare Function GetWindowThreadProcessId Lib "user32" _
          (ByVal hWnd As Long, lpdwProcessId As Long) As Long

  ' SendMessage function sends the specified message to a window or
  ' windows. The function calls the window procedure for the specified
  ' window and does not return until the window procedure has processed
  ' the message.
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
          (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
          ByVal lParam As Any) As Long

  ' GetProcessImageFileName function returns the path in device form,
  ' rather than drive letters.
  Private Declare Function GetProcessImageFileName Lib "psapi.dll" _
          Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, _
          ByVal lpImageFileName As String, ByVal nSize As Long) As Long

  '================= Enumerating Parent and Child windows ====================
  ' The EnumWindows() function enumerates all top-level windows on the screen
  ' by passing the handle of each window, in turn, to an application-defined
  ' callback function. EnumWindows() continues until the last top-level window
  ' is enumerated or the callback function returns FALSE.
  Private Declare Function EnumWindows Lib "user32" _
          (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

  ' The EnumChildWindows function enumerates the child windows that belong to
  ' the specified parent window by passing the handle of each child window, in
  ' turn, to an application-defined callback function. EnumChildWindows
  ' continues until the last child window is enumerated or the callback
  ' function returns FALSE.  If a child window has created child windows of
  ' its own, this function enumerates those windows as well.
  Private Declare Function EnumChildWindows Lib "user32" _
          (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, _
          ByVal lParam As Long) As Boolean

  ' The GetWindowText() function copies the text of the specified window's
  ' (Parent) title bar (if it has one) into a buffer. If the specified window
  ' is a control, the text of the control is copied.
  Private Declare Function GetWindowText Lib "user32" _
          Alias "GetWindowTextA" (ByVal hWnd As Long, _
          ByVal lpString As String, ByVal cch As Long) As Long

  ' The GetClassName() function retrieves the name of the class to which the
  ' specified window belongs.
  Private Declare Function GetClassName Lib "user32" _
          Alias "GetClassNameA" (ByVal hWnd As Long, _
          ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  '===========================================================================

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Boolean)
'                    |  |     |----- Variable subname
'                    - --- -------------
' Naming standard:   m bln FindSpecific
' Variable name:     mblnFindSpecific
' ***************************************************************************
  Private mblnFindSpecific    As Boolean
  Private mblnGetChildWindows As Boolean
  Private mlngHwnd            As Long
  Private mlngIndex           As Long
  Private mlngArraySize       As Long
  Private mstrComputer        As String
  Private mstrParentName      As String   ' Name of parent process
  Private mstrSearchItem      As String
  Private mastrTemp()         As String   ' Collection of process names

' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       StopStubbornPgms
'
' Description:   This routine is used to stop processes that are known for
'                hanging when shutting down a PC.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Feb-2014  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Sub StopStubbornPgms()

    Dim lngIndex   As Long
    Dim astrData() As String

    Const APPL_CNT As Long = 20

    ' Verify most of the active applications
    ' are deactivated.  Makes for an easier
    ' shutdown process.

    Erase astrData()           ' Always start with empty arrays
    ReDim astrData(APPL_CNT)   ' Size temp array

    ' Preload
    For lngIndex = 0 To UBound(astrData) - 1
        astrData(lngIndex) = ""
    Next lngIndex

    ' Ex:  Load array with names of known applications
    '      that sometimes hang during shutdown process
    astrData(0) = "winword.exe"    ' MS Word (Document writer)
    astrData(1) = "excel.exe"      ' MS Excel (Spreadsheet)
    astrData(2) = "outlook.exe"    ' MS Outlook (Mail provider - client)
    astrData(3) = "msaccess.exe"   ' MS Access (Database)
    astrData(4) = "msqry32.exe"    ' MS SQL (Database)
    astrData(5) = "powerpnt.exe"   ' MS Powerpoint (Presentation)
    astrData(6) = "itunes.exe"     ' Apple software (Music, movies, etc.)

    ' Loop thru array
    For lngIndex = 0 To UBound(astrData) - 1

        DoEvents
        If Len(astrData(lngIndex)) > 0 Then

            ' Search for process by executable name
            StopProcessByName astrData(lngIndex)
            DoEvents
        Else
            Exit For   ' exit FOR..NEXT loop
        End If

    Next lngIndex

    Erase astrData()   ' Always empty arrays when not needed

End Sub

' ***************************************************************************
' Routine:       FindProcessByHandle
'
' Description:   This routine will perform a search of all active processes,
'                either hidden, minimized, or displayed.  If found, the
'                procerss handle will be returned.  Optionally, the process
'                executable name is also returned.
'
' Parameters:    lngHwnd - Specific process handle to search for
'                strExeName - Optional - Full executable name
'                          (ex: "OUTLOOK.EXE")
'
' Returns:       Returns first process handle and full executable name
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function FindProcessByHandle(ByVal lngHwnd As Long, _
                           Optional ByRef strExeName As String = vbNullString) As Long

    ' Called by StopProcessByHandle()

    Dim objWMI       As Object
    Dim objProcess   As Object
    Dim objProcesses As Object

    FindProcessByHandle = 0    ' Preset to not found
    Call CaptureComputerName   ' Capture this computer name

    Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & mstrComputer & "/root/cimv2")
    Set objProcesses = objWMI.ExecQuery("Select * from Win32_Process")

    For Each objProcess In objProcesses

        ' Look for matching handle value
        With objProcess
            If lngHwnd = .Handle Then
                strExeName = .Name              ' Found name of process
                FindProcessByHandle = .Handle   ' Return handle number
                Exit For
            End If
        End With

    Next objProcess

    ' Free objects from memory
    Set objWMI = Nothing
    Set objProcess = Nothing
    Set objProcesses = Nothing

End Function

' ***************************************************************************
' Routine:       FindHandleByExe
'
' Description:   Find handle to a specific executable
'
' Parameters:    strExeName - Name of executable (ex:  "iexplore.exe")
'
' Returns:       If found, returns numeric handle value else zero
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function FindHandleByExe(ByVal strExeName As String) As Long

    Dim objWMI       As Object
    Dim objProcess   As Object
    Dim objProcesses As Object

    FindHandleByExe = 0        ' Preset to not found
    Call CaptureComputerName   ' Capture this computer name

    Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & mstrComputer & "/root/cimv2")
    Set objProcesses = objWMI.ExecQuery("Select * from Win32_Process")

    For Each objProcess In objProcesses

        ' Look for executable name
        With objProcess

            ' Verify data is not Null
            If Len(.Name) > 0 Then

                ' Compare to name to be matched
                If InStr(1, .Name, strExeName, vbTextCompare) > 0 Then
                    FindHandleByExe = .Handle   ' Successful finish
                    Exit For
                End If
            End If
        End With

    Next objProcess

    ' Free objects from memory
    Set objWMI = Nothing
    Set objProcess = Nothing
    Set objProcesses = Nothing

End Function

' ***************************************************************************
' Routine:       FindProcessByCaption
'
' Description:   This routine will perform a search of all active processes,
'                either hidden, minimized, or displayed.
'
' Parameters:    strCaption - partial/full name of caption title
'
' Returns:       Returns process handle
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Feb-2014  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function FindProcessByCaption(ByVal strCaption As String) As Long

    ' Called by StopProcessByCaption()

    Dim lngHwnd        As Long
    Dim blnFoundit     As Boolean
    Dim strFullCaption As String

    FindProcessByCaption = 0                           ' Preload for not found
    blnFoundit = False                                 ' Preset flag to FALSE
    lngHwnd = FindWindow(vbNullString, vbNullString)   ' Find first window handle

    Do While lngHwnd <> 0

        strFullCaption = GetFullCaption(lngHwnd)   ' Capture full window caption

        ' See if our data string is within this caption
        If InStr(1, strFullCaption, strCaption, vbTextCompare) > 0 Then
            FindProcessByCaption = lngHwnd   ' Return handle
            blnFoundit = True                ' Set flag to TRUE
            Exit Do                          ' Exit Do..Loop
        End If

        lngHwnd = GetWindow(lngHwnd, GW_HWNDNEXT)   ' Get next window
    Loop

    If Not blnFoundit Then
        If lngHwnd <> 0 Then
            Call CloseHandle(lngHwnd)    ' Always close handles when not needed
        End If
    End If

End Function

' ***************************************************************************
' Routine:       StopProcessByHandle
'
' Description:   This routine is used to stop a specific process if the
'                process handle is known.
'
' Parameters:    lngHwnd - Unique handle designating process to be closed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Feb-2014  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function StopProcessByHandle(ByVal lngHwnd As Long) As Boolean

    Dim strExeName   As String
    Dim lngProcessID As Long

    StopProcessByHandle = False   ' Preset to FALSE

    ' Search for a specific process handle
    lngProcessID = FindProcessByHandle(lngHwnd, strExeName)

    ' If found then kill all occurances
    If lngProcessID <> 0 Then

        If Len(strExeName) > 0 Then
            StopProcessByHandle = StopProcessByName(strExeName)
        End If
    End If

    ' Always close handles when not needed
    If lngProcessID <> 0 Then
        Call CloseHandle(lngProcessID)
    End If

End Function

' ***************************************************************************
' Routine:       StopProcessByCaption
'
' Description:   This routine will perform a search of all active processes,
'                either hidden, minimized, or displayed.  If found, the
'                parent procss handle (if any) will be used to get the
'                name of the process executable.  This executable will be
'                closed.
'
' Parameters:    strCaption - Full or partial caption data to search for
'
' Returns:       TRUE if successful else FALSE
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Feb-2014  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function StopProcessByCaption(ByVal strCaption As String) As Boolean

    Dim strExeName   As String
    Dim lngHwnd      As Long
    Dim lngProcessID As Long

    StopProcessByCaption = False       ' Preset to FALSE
    strCaption = TrimStr(strCaption)   ' Remove unwanted trailing characters

    If Len(strCaption) > 0 Then

        ' Search for process by caption title
        lngHwnd = FindProcessByCaption(strCaption)

        If lngHwnd <> 0 Then

            ' Get parent application handle
            lngProcessID = GetParentProcessID(lngHwnd)

            ' Is this the parent process?
            If lngProcessID <> 0 Then
                lngHwnd = lngProcessID
            End If

            ' Get name of parent executable
            strExeName = ExeNameFromProcID(lngHwnd)

            ' If found then kill all occurances
            If Len(strExeName) > 0 Then
                StopProcessByCaption = StopProcessByName(strExeName)
            End If

        End If
    End If

    ' Always close handles when not needed
    If lngProcessID > 0 Then
        Call CloseHandle(lngProcessID)
    End If

    If lngHwnd > 0 Then
        Call CloseHandle(lngHwnd)
    End If

End Function

' ***************************************************************************
' Routine:       StopProcessByName
'
' Description:   The following Function terminates all processes using
'                the name of a particular executable (ex: winword.exe)
'                This has the same effect as pressing the "End Task"
'                button in Task Mananger.
'
' Reference:     SyS_V|rUS 27-Nov-2003
'                http://www.rohitab.com/discuss/topic/5962-terminate-process-in-vb6/
'
' Parameters:    strExeName - Name of the executable to be closed
'
' Returns:       True if successful
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Feb-2014  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function StopProcessByName(ByVal strExeName As String) As Boolean

    ' Called by StopStubbornPgms()
    '           StopProcessByHandle()
    '           StopProcessByCaption()

    Dim strTmpName      As String
    Dim lngHwnd         As Long
    Dim lngCount        As Long
    Dim lngProcessID    As Long
    Dim lngProcessFound As Long
    Dim typProcess      As PROCESSENTRY32

    On Error GoTo StopProcessByName_Error

    lngCount = 0
    ZeroMemory typProcess, Len(typProcess)   ' Clear type structure
    typProcess.dwSize = Len(typProcess)      ' Initialize type structure

    lngHwnd = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    lngProcessFound = Process32First(lngHwnd, typProcess)

    Do While lngProcessFound

        strTmpName = LCase$(TrimStr(typProcess.szExeFile))   ' Remove unwanted data

        ' Is this same executable
        If InStr(1, strTmpName, strExeName, vbTextCompare) > 0 Then

            ' Verify process is active
            lngProcessID = OpenProcess(1&, -1&, typProcess.th32ProcessID)

            ' Do we have a handle for this process
            If lngProcessID <> 0 Then

                ' Force process to close
                TerminateProcess lngProcessID, 0&
            End If

            Call CloseHandle(lngProcessID)   ' Always close handles when not needed
            StopProcessByName = True         ' Set flag to TRUE
            lngCount = lngCount + 1          ' Increment occurance counter

        End If

        ZeroMemory typProcess, Len(typProcess)   ' Clear type structure
        typProcess.dwSize = Len(typProcess)      ' Initialize type structure

        ' Find next occurance of this process
        lngProcessFound = Process32Next(lngHwnd, typProcess)
        DoEvents

    Loop

StopProcessByName_CleanUp:
    If lngHwnd <> 0 Then
        Call CloseHandle(lngHwnd)   ' Always close handles when not needed
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

StopProcessByName_Error:
    Err.Clear                          ' Reset any error codes
    StopProcessByName = False          ' Set return flag to FALSE
    Resume StopProcessByName_CleanUp

End Function

' ***************************************************************************
' Routine:       GetParentChildWindows
'
' Description:   Locates active (Parent) and Child windows.
'
' Parameters:    astrData() - String array to hold windows data
'                strSpecificCaption - Optional - If looking for a specific window
'                    then this will contain a complete or partial window
'                    caption name.  Default = vbNullString
'                lngHwnd - Optional - If looking for a specific window
'                    then this will contain a value that matches the
'                    specific windows handle.  Default = 0
'                blnGetChildWindows - Optional - Capture all Child windows
'                    associated with a Parent window.  Default = TRUE
'
'                ex:   ' Return an array of all window processes and handles
'                      Call GetParentChildWindows(astrData())
'
' Returns:       An array of Window parent and child windows along
'                with any appropriate handle values.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 25-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to capture Parent windows only
' ***************************************************************************
Public Function GetParentChildWindows(ByRef astrData() As String, _
                             Optional ByVal strSpecificCaption As String = vbNullString, _
                             Optional ByVal lngHwnd As Long = 0&, _
                             Optional ByVal blnGetChildWindows As Boolean = True) As Boolean

    Dim lngIdx As Long

    On Error GoTo GetParentChildWindows_CleanUp

    GetParentChildWindows = False              ' Preset flag for failure
    mblnGetChildWindows = blnGetChildWindows   ' True or False (User choice)
    ReDim astrData(0)                          ' Preload with one entry
    mlngArraySize = ARRY_SIZE * 5              ' Starting array size

    If (Len(strSpecificCaption) > 0) Or (lngHwnd > 0) Then
        mblnFindSpecific = True
        mstrSearchItem = strSpecificCaption
        mlngHwnd = lngHwnd
    Else
        mblnFindSpecific = False
        mstrSearchItem = ""
        mlngHwnd = 0
    End If

    mlngIndex = 0                    ' Set array index
    ReDim mastrTemp(mlngArraySize)   ' Size return array

    ' Enumerate all active windows using
    ' the AddressOf callback function
    EnumWindows AddressOf GetParentWindows, 0&

    If mlngIndex > 0 Then

        ReDim astrData(mlngIndex)   ' Resize return array

        ' Return collected data
        For lngIdx = 0 To (mlngIndex - 1)
            astrData(lngIdx) = mastrTemp(lngIdx)
        Next lngIdx

        GetParentChildWindows = True   ' Set flag for success

    End If

GetParentChildWindows_CleanUp:
    If Err.Number <> 0 Then
        Err.Clear     ' Clear error code, if any
    End If

    Erase mastrTemp()             ' Empty temp array
    mblnGetChildWindows = False   ' Reset to FALSE
    On Error GoTo 0               ' Nullify this error trap

End Function


' ***************************************************************************
' ****               Internal procedures & functions                     ****
' ***************************************************************************

Private Function GetFullCaption(ByVal lngHwnd As Long) As String

    ' Called by FindProcessByCaption()

    Dim lngLength  As Long
    Dim strCaption As String

    strCaption = ""
    lngLength = SendMessage(lngHwnd, WM_GETTEXTLENGTH, 0&, 0&)

    If lngLength > 0 Then
        strCaption = String$(lngLength, 0&)
        SendMessage lngHwnd, WM_GETTEXT, lngLength + 1, strCaption
        GetFullCaption = strCaption
    End If

End Function

Private Function GetParentProcessID(ByVal lngHwnd As Long) As Long

    ' Called by StopProcessByCaption()
    '           FindProcessByHandle()

    Dim lngProcessID As Long

    Call GetWindowThreadProcessId(lngHwnd, lngProcessID)
    GetParentProcessID = lngProcessID

End Function

Private Function ExeNameFromProcID(ByVal lngProcessID As Long) As String

    ' Called by StopProcessByCaption()
    '           FindProcessByHandle()

    Dim lngHwnd    As Long
    Dim strExeName As String

    strExeName = Space$(MAX_SIZE)
    lngHwnd = OpenProcess(PROCESS_QUERY_INFO Or PROCESS_VM_READ, 0, lngProcessID)

    If lngHwnd Then

        If GetProcessImageFileName(lngHwnd, strExeName, Len(strExeName)) <> 0 Then
            strExeName = TrimStr(strExeName)
            ExeNameFromProcID = GetFilename(strExeName)
        End If

        Call CloseHandle(lngHwnd)

    End If

End Function

Private Function GetFilename(ByVal strPath As String) As String

    ' Called by ExeNameFromProcID()

    Dim lngPointer As Long

    ' Find last backslash in string
    lngPointer = InStrRev(strPath, "\")

    If lngPointer > 0 Then
        GetFilename = Mid$(strPath, lngPointer + 1)
    Else
        GetFilename = strPath
    End If

End Function

Private Sub CaptureComputerName()

    ' Called by FindProcessByHandle()
    '           FindHandleByExe()

    Dim objWMI As Object

    ' See if we already have a name
    If Len(TrimStr(mstrComputer)) > 0 Then
        Exit Sub
    End If

    ' Capture name of this computer
    Set objWMI = CreateObject("Wscript.Network")   ' Instantiate object
    mstrComputer = objWMI.ComputerName             ' Capture this computer name
    Set objWMI = Nothing                           ' Free object from memory

    If Len(TrimStr(mstrComputer)) = 0 Then
        mstrComputer = "."   ' Any computer (Period equal wildcard)
    End If

End Sub

' ***************************************************************************
' Routine:       GetParentWindows
'
' Description:   Locates all active (parent) windows.  Then adds the parent
'                and the class process handle values to the accumulator.
'                Then makes a call to another routine to search for any
'                associated child windows.
'
'                         +---------------------------------- Parent process handle (number)
'                         |       +-------------------------- Parent process name (string)
'                ex:  "1770982|DiscDataWipe - modMain (Code)"
'                             +---- Data separator
'
' Parameters:    lngHwnd - Value of current active process.  Loaded within
'                          this routine.
'                lngParm - Special parameters (Not used but required)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 25-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to capture Parent windows only
' ***************************************************************************
Private Function GetParentWindows(ByVal lngHwnd As Long, _
                                  ByVal lngParm As Long) As Long

    ' Called by GetParentChildWindows()

    Dim strParent As String

    strParent = Space$(MAX_SIZE)                 ' Preload with spaces (not nulls)
    GetWindowText lngHwnd, strParent, MAX_SIZE   ' API call to get parent name
    mstrParentName = TrimStr(strParent)          ' Capture full name of parent process

    If lngHwnd > 0 Then

        If mblnFindSpecific Then

            ' Look for a specific process
            If mlngHwnd > 0 Then

                If mlngHwnd = lngHwnd Then

                    '                         +--------------------------- Parent process handle (number)
                    '                         |       +------------------- Data separator
                    '                         |       |        +---------- Parent process name (string)
                    mastrTemp(mlngIndex) = lngHwnd & "|" & mstrParentName

                    mlngIndex = mlngIndex + 1   ' Increment index pointer

                    ' Update array size if needed
                    If mlngIndex >= mlngArraySize Then
                        DoEvents                                    ' Slow processing down to allow update
                        mlngArraySize = mlngArraySize + ARRY_SIZE   ' Increase number of array entries
                        ReDim Preserve mastrTemp(mlngArraySize)     ' Update array size w/o losing data
                    End If

                    ' Make API call to find all child windows
                    ' associated with this window using the
                    ' AddressOf callback function. Parameters
                    ' passed are the current active process
                    ' handle from above and name of routine
                    ' that will loop thru all child windows
                    If mblnGetChildWindows Then
                        EnumChildWindows lngHwnd, AddressOf GetChildWindows, 0&
                    End If

                    GetParentWindows = 0   ' No more searching
                End If

            ElseIf Len(mstrSearchItem) > 0 Then

                If InStr(1, strParent, mstrSearchItem, vbTextCompare) > 0 Then

                    mlngHwnd = lngHwnd   ' Capture process handle

                    '                         +--------------------------- Parent process handle (number)
                    '                         |       +------------------- Data separator
                    '                         |       |        +---------- Parent process name (string)
                    mastrTemp(mlngIndex) = lngHwnd & "|" & mstrParentName

                    mlngIndex = mlngIndex + 1   ' Increment index pointer

                    ' Update array size if needed
                    If mlngIndex >= mlngArraySize Then
                        DoEvents                                    ' Slow processing down to allow update
                        mlngArraySize = mlngArraySize + ARRY_SIZE   ' Increase number of array entries
                        ReDim Preserve mastrTemp(mlngArraySize)     ' Update array size w/o losing data
                    End If

                    ' Make API call to find all child windows
                    ' associated with this window using the
                    ' AddressOf callback function. Parameters
                    ' passed are the current active process
                    ' handle from above and name of routine
                    ' that will loop thru all child windows
                    If mblnGetChildWindows Then
                        EnumChildWindows lngHwnd, AddressOf GetChildWindows, 0&
                    End If

                    GetParentWindows = 0   ' No more searching
                End If
            End If

        Else   ' Look for all processes

            mlngHwnd = lngHwnd   ' Capture process handle

            '                         +--------------------------- Parent process handle (number)
            '                         |       +------------------- Data separator
            '                         |       |        +---------- Parent process name (string)
            mastrTemp(mlngIndex) = lngHwnd & "|" & mstrParentName

            mlngIndex = mlngIndex + 1   ' Increment index pointer

            ' Update array size if needed
            If mlngIndex >= mlngArraySize Then
                DoEvents                                    ' Slow processing down to allow update
                mlngArraySize = mlngArraySize + ARRY_SIZE   ' Increase number of array entries
                ReDim Preserve mastrTemp(mlngArraySize)     ' Update array size w/o losing data
            End If

            ' Make API call to find all child windows
            ' associated with this window using the
            ' AddressOf callback function. Parameters
            ' passed are the current active process
            ' handle from above and name of routine
            ' that will loop thru all child windows
            If mblnGetChildWindows Then
                EnumChildWindows lngHwnd, AddressOf GetChildWindows, 0&
            End If

            GetParentWindows = -1   ' Look for next active (parent) process
        End If
    End If

    CloseHandle lngHwnd   ' Always release handle

End Function

' ***************************************************************************
' Routine:       GetChildWindows
'
' Description:   Determines if there is a child process. If yes, then
'                searches for any additional child windows.
'
'                         +--------------------------------------------- Parent process handle (number)
'                         |       +------------------------------------- Parent process name (string)
'                         |       |         +--------------------------- Child process handle (number)
'                         |       |         |       +------------------- Child process name (string)
'                ex:  "1901938|Jump List|1770898|DesktopDestinationList"
'                             +_________+_______+
'                                       |
'                                       +---- Data separators
'
' Parameters:    lngHwnd - Value of current child process.  Loaded
'                          within this routine.
'                lngParm - Special parameter (Not used but required)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function GetChildWindows(ByVal lngHwnd As Long, _
                                 ByVal lngParm As Long) As Long

    ' Called by GetParentWindows()

    Dim strChildName As String

    strChildName = Space$(MAX_SIZE)   ' Preload with spaces (not nulls)

    ' Get Child window name
    If GetClassName(lngHwnd, strChildName, MAX_SIZE) <> 0 Then

        strChildName = TrimStr(strChildName)   ' Remove unwanted spaces

        '                          +---------------------------------------------------------------- Parent process handle (number)
        '                          |                +----------------------------------------------- Parent process name (string)
        '                          |                |                     +------------------------- Child process handle (number)
        '                          |                |                     |               +--------- Child process name (string)
        mastrTemp(mlngIndex) = mlngHwnd & "|" & mstrParentName & "|" & lngHwnd & "|" & strChildName
        '                                  +______________________+_______________+
        '                                                         |
        '                                                         +---- Data separators
        mlngIndex = mlngIndex + 1   ' Increment index pointer

        ' Update array size if needed
        If mlngIndex >= mlngArraySize Then
            DoEvents                                    ' Slow processing down to allow update
            mlngArraySize = mlngArraySize + ARRY_SIZE   ' Increase number of array entries
            ReDim Preserve mastrTemp(mlngArraySize)     ' Update array size w/o losing data
        End If

    End If

    CloseHandle lngHwnd    ' Always release handle
    GetChildWindows = -1   ' Look for next child process

End Function

