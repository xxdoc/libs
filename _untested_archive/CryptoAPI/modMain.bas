Attribute VB_Name = "modMain"
' ***************************************************************************
' Module:        modMain
'
' Description:   This is a generic module I use to start and stop an
'                application
'
' IMPORTANT:     If this application does not execute as it should, then
'                right mouse click the application name and select
'                'Run as Administrator' or 'Properties>>Compatibility' tab
'                and check the box to 'Run this program as an administrator'.
'
'                With Windows 8 and newer, Microsoft will split Admin
'                users rights via User Access Control (UAC) to where the
'                user will have to use the option to 'Run as Administrator'.
'                The only other option will be to turn off UAC which is not
'                recommended because everyone and any application will have
'                authority to make change to your machine without any
'                interference.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced FileExists() and PathExists() routines with
'              IsPathValid() routine.
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Deleted RemoveTrailingNulls() routine from this module.
'              - Changed call to RemoveTrailingNulls() to TrimStr module
'                due to speed and accuracy.
' 15-May-2015  Kenneth Ives  kenaso@tx.rr.com
'              Added FindDefaultBrowser() routine
' 08-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Renamed routine FindDefaultBrowser() to FindDefaultPath().
'              - Modified FindDefaultPath() routine to be more generic.
' 28-Mar-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added AdjustPrivileges() routine
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Global constants
' ***************************************************************************
  Public Const SUPPORT_EMAIL As String = "kenaso@tx.rr.com"
  Public Const PGM_NAME      As String = "CryptoAPI Demo"
  Public Const MAX_SIZE      As Long = 260

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME          As String = "modMain"
  Private Const SE_DEBUG_NAME        As String = "SeDebugPrivilege"
  Private Const ERROR_ALREADY_EXISTS As Long = 183&

  ' Used in DisableX() routine
  Private Const MF_BYPOSITION As Long = &H400
  Private Const MF_REMOVE     As Long = &H1000

  ' Used in ShellAndWait() routine
  Private Const STATUS_PENDING            As Long = &H103&
  Private Const PROCESS_QUERY_INFORMATION As Long = &H400

  ' Used in FindDefaultPath() routine
  Private Const ERROR_SUCCESS      As Long = 0&
  Private Const REG_SZ             As Long = 1&
  Private Const REG_EXPAND_SZ      As Long = 2&
  Private Const KEY_READ           As Long = &H20119

  ' Used to initiate manifest file
  ' Set of bit flags that indicate which common control classes
  ' will be loaded.  The dwICC value of INIT_COMMON_CTRLS can
  ' be a combination of the following:
  Private Const ICC_ANIMATE_CLASS      As Long = &H80&     ' Load animate control class
  Private Const ICC_BAR_CLASSES        As Long = &H4&      ' Load toolbar, status bar, trackbar, tooltip control classes
  Private Const ICC_COOL_CLASSES       As Long = &H400&    ' Load rebar control class
  Private Const ICC_DATE_CLASSES       As Long = &H100&    ' Load date and time picker control class
  Private Const ICC_HOTKEY_CLASS       As Long = &H40&     ' Load hot key control class
  Private Const ICC_INTERNET_CLASSES   As Long = &H800&    ' Load IP address class
  Private Const ICC_LINK_CLASS         As Long = &H8000&   ' Load a hyperlink control class. Must have trailing ampersand.
  Private Const ICC_LISTVIEW_CLASSES   As Long = &H1&      ' Load list-view and header control classes
  Private Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&   ' Load a native font control class
  Private Const ICC_PAGESCROLLER_CLASS As Long = &H1000&   ' Load pager control class
  Private Const ICC_PROGRESS_CLASS     As Long = &H20&     ' Load progress bar control class
  Private Const ICC_STANDARD_CLASSES   As Long = &H4000&   ' Load user controls that include button, edit, static, listbox,
                                                           '      combobox, scrollbar
  Private Const ICC_TREEVIEW_CLASSES   As Long = &H2&      ' Load tree-view and tooltip control classes
  Private Const ICC_TAB_CLASSES        As Long = &H8&      ' Load tab and tooltip control classes
  Private Const ICC_UPDOWN_CLASS       As Long = &H10&     ' Load up-down control class
  Private Const ICC_USEREX_CLASSES     As Long = &H200&    ' Load ComboBoxEx class
  Private Const ICC_WIN95_CLASSES      As Long = &HFF&     ' Load animate control, header, hot key, list-view, progress bar,
                                                           '      status bar, tab, tooltip, toolbar, trackbar, tree-view,
                                                           '      and up-down control classes

  ' All bit flags combined. Total value = &HFFFF& (65535)
  Private Const ICC_ALL_CLASSES As Long = ICC_ANIMATE_CLASS Or ICC_BAR_CLASSES Or ICC_COOL_CLASSES Or _
                                          ICC_DATE_CLASSES Or ICC_HOTKEY_CLASS Or ICC_INTERNET_CLASSES Or _
                                          ICC_LINK_CLASS Or ICC_LISTVIEW_CLASSES Or ICC_NATIVEFNTCTL_CLASS Or _
                                          ICC_PAGESCROLLER_CLASS Or ICC_PROGRESS_CLASS Or ICC_STANDARD_CLASSES Or _
                                          ICC_TREEVIEW_CLASSES Or ICC_TAB_CLASSES Or ICC_UPDOWN_CLASS Or _
                                          ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES

' ***************************************************************************
' Type structures
' ***************************************************************************
  Private Type INIT_COMMON_CTRLS
      dwSize As Long   ' size of this structure
      dwICC  As Long   ' flags indicating which classes to be initialized
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Private Declare Function GetTickCount Lib "kernel32" () As Long

  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long

  '  OpenProcess function returns a handle of an existing process object.
  ' If the function succeeds, the return value is an open handle of the
  ' specified process.
  Private Declare Function OpenProcess Lib "kernel32" _
          (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
          ByVal dwProcessId As Long) As Long

  ' The GetCurrentProcess function returns a pseudohandle for the current
  ' process. A pseudohandle is a special constant that is interpreted as
  ' the current process handle. The calling process can use this handle to
  ' specify its own process whenever a process handle is required. The
  ' pseudohandle need not be closed when it is no longer needed.
  Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

  ' The GetExitCodeProcess function retrieves the termination status of the
  ' specified process. If the function succeeds, the return value is nonzero.
  Private Declare Function GetExitCodeProcess Lib "kernel32" _
          (ByVal hProcess As Long, lpExitCode As Long) As Long

  ' ExitProcess function ends a process and all its threads
  ' ex:     ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
  Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

  ' The CreateMutex function creates a named or unnamed mutex object.  Used
  ' to determine if an application is active.
  Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" _
          (lpMutexAttributes As Any, ByVal bInitialOwner As Long, _
          ByVal lpName As String) As Long

  ' This function releases ownership of the specified mutex object.
  ' Finished with the search.
  Private Declare Function ReleaseMutex Lib "kernel32" _
          (ByVal hMutex As Long) As Long

  ' GetDesktopWindow function retrieves a handle to the desktop window.
  ' The desktop window covers the entire screen. The desktop window is
  ' the area on top of which other windows are painted. The return
  ' value is a handle to the desktop window.
  Private Declare Function GetDesktopWindow Lib "user32" () As Long

  ' The ShellExecute function opens or prints a specified file.  The file
  ' can be an executable file or a document file.
  Private Declare Function ShellExecute Lib "shell32.dll" _
          Alias "ShellExecuteA" (ByVal hWnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, _
          ByVal lpParameters As String, ByVal lpDirectory As String, _
          ByVal nShowCmd As Long) As Long

  ' Always close a handle if not being used
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long

  ' Truncates a path to fit within a certain number of characters by replacing
  ' path components with ellipses.
  Private Declare Function PathCompactPathEx Lib "shlwapi.dll" _
          Alias "PathCompactPathExA" _
          (ByVal pszOut As String, ByVal pszSrc As String, _
          ByVal cchMax As Long, ByVal dwFlags As Long) As Long

  ' RegOpenKeyEx function opens the specified key.  Returns zero if successful.
  Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
          (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
          ByVal samDesired As Long, phkResult As Long) As Long

  ' RegQueryValueEx function retrieves the type and data for a specified value
  ' name associated with an open registry key.  Returns zero if successful.
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
          (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
          lpType As Long, lpData As String, lpcbData As Long) As Long

  ' RegCloseKey function releases the handle of the specified key.
  Private Declare Function RegCloseKey Lib "advapi32.dll" _
          (ByVal hkey As Long) As Long

  ' ========= Used to DisableX on form ======================================
  ' The DrawMenuBar function redraws the menu bar of the specified window.
  ' If the menu bar changes after Windows has created the window, this
  ' function must be called to draw the changed menu bar.  If the function
  ' fails, the return value is zero.
  Private Declare Function DrawMenuBar Lib "user32" _
          (ByVal hWnd As Long) As Long

  ' The GetMenuItemCount function determines the number of items in the
  ' specified menu.  If the function fails, the return value is -1.
  Private Declare Function GetMenuItemCount Lib "user32" _
          (ByVal hMenu As Long) As Long

  ' The GetSystemMenu function allows the application to access the window
  ' menu (also known as the System menu or the Control menu) for copying
  ' and modifying.  If the bRevert parameter is FALSE (0&), the return
  ' value is the handle of a copy of the window menu.  If the function
  ' fails, the return value is zero.
  Private Declare Function GetSystemMenu Lib "user32" _
          (ByVal hWnd As Long, ByVal bRevert As Long) As Long

  ' The RemoveMenu function deletes a menu item from the specified menu.
  ' If the menu item opens a drop-down menu or submenu, RemoveMenu does
  ' not destroy the menu or its handle, allowing the menu to be reused.
  ' Before this function is called, the GetSubMenu function should retrieve
  ' the handle of the drop-down menu or submenu.  If the function fails,
  ' the return value is zero.
  Private Declare Function RemoveMenu Lib "user32" _
          (ByVal hMenu As Long, ByVal nPosition As Long, _
          ByVal wFlags As Long) As Long
  ' =========================================================================

  ' ========= Initialize Manifest file ======================================
  ' Initializes specific common controls classes from the common control
  ' dynamic-link library. Returns TRUE (non-zero) if successful, or FALSE
  ' otherwise. Began being exported with Comctl32.dll version 4.7
  ' (IE3.0 & later).
  Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
          (iccex As INIT_COMMON_CTRLS) As Boolean

  ' Initializes the entire common control dynamic-link library. Exported by
  ' all versions of Comctl32.dll.
  Private Declare Sub InitCommonControls Lib "comctl32" ()
  ' =========================================================================

' ***************************************************************************
' API Declarations (Public)
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function
  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
' Variable name:     gstrVersion
' Naming standard:   g str Version
'                    - --- ---------
'                    |  |    |______ Variable subname
'                    |  |___________ Data type (String)
'                    |______________ Global level designator
'
' ***************************************************************************
  Public gblnWin8or81   As Boolean
  Public gstrVersion    As String
  Public gstrOperSystem As String

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Boolean)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m bln VB_IDE
' Variable name:     mblnVB_IDE
' ***************************************************************************
  Private mblnVB_IDE     As Boolean
  Private mobjPrivileges As cPrivileges


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Main
'
' Description:   This is a generic routine to start an application
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Main()

    Dim lngMajorVer     As Long
    Dim lngMinorVer     As Long
    Dim blnWinXP_SP3    As Boolean
    Dim blnOperSystem64 As Boolean
    Dim objOperSys      As cOperSystem
    Dim objManifest     As cManifest

    Const ROUTINE_NAME  As String = "Main"

    On Error Resume Next
    ChDrive App.Path
    ChDir App.Path
    On Error GoTo 0

    On Error GoTo Main_Error

    ' See if there is another instance of this program
    ' running.  The parameter being passed is the name
    ' of this executable without the EXE extension.
    If Not AlreadyRunning(App.EXEName) Then

        gstrVersion = PGM_NAME & " v" & App.Major & "." & App.Minor & "." & App.Revision
        gblnStopProcessing = False        ' Preset global stop flag

        Load frmSplash   ' Load splash screen
        Wait 1000

        ' Instantiate class objects
        Set objOperSys = New cOperSystem
        Set mobjPrivileges = New cPrivileges

        With objOperSys
            ' Capture information about this operating system
            gstrOperSystem = .VersionName & vbNewLine & TrimStr("Ver " & .VersionData) & " " & _
                             .ProcessArchitecture & " (" & .BaseBitStructure & ")"
            gblnWin8or81 = .bCaptionIsCentered                ' If TRUE then this is Win 8 or 8.1
            lngMajorVer = CLng(.MajorVersion)                 ' Capture OS major version
            lngMinorVer = CLng(.MinorVersion)                 ' Capture OS minor version
            blnOperSystem64 = .bOperSystem64                  ' Is OS 32/64-bit
            blnWinXP_SP3 = .bWinXP_SP3orNewer                 ' If XP, is SP3 installed
            mobjPrivileges.bWindows7 = .bWindows7             ' Is OS Windows 7
            mobjPrivileges.bWindows8 = .bCaptionIsCentered    ' If TRUE then this is Win 8 or 8.1

            ' Is OS Windows 10 (Beta or Retail)
            If .bWindows10_Beta Or .bWindows10 Then
                mobjPrivileges.bWindows10 = True
            End If
        End With

        ' Test for correct operating system
        Select Case lngMajorVer

               Case Is >= 6   ' Windows Vista or newer
                    ' All is good

               Case 5   ' Windows XP
                    If lngMinorVer = 2 Then
                        If blnOperSystem64 Then
                            If blnWinXP_SP3 Then
                                ' All is good
                            Else
                                InfoMsg "Recommend that you install XP Service Pack 3" & vbNewLine & _
                                        "or this application may not function properly.", , "Warning"
                            End If
                        Else
                            InfoMsg "This application may not function properly" & vbNewLine & _
                                    "under Windows XP 32-bit operating system.", , "Warning"
                        End If

                    Else
                        InfoMsg "This application may not function properly" & vbNewLine & _
                                "under this version of Windows XP.", , "Warning"
                    End If

               Case Else   ' Earlier versions of Windows
                    InfoMsg "This application does not support" & vbNewLine & _
                            "this version of Windows.", , "Warning"

                    Unload frmSplash    ' Unload splash screen
                    GoTo Main_CleanUp   ' Close this application
        End Select

        ' Temporarily update user privileges
        With mobjPrivileges
            .bVB_IDE = mblnVB_IDE
            '                  +-------------------------- Enable a privilege (True\False)
            '                  |        +----------------- Name of temporary privilege
            '                  |        |           +----- Test if user has Admin authority
            '                  |        |           |
            .CheckPrivileges True, SE_DEBUG_NAME, True
        End With

        ' Create/Update manifest file
        Set objManifest = New cManifest
        With objManifest
            .MajorVersion = lngMajorVer           ' OS major version
            .MinorVersion = lngMinorVer           ' OS minor version
            .bVB_IDE = mblnVB_IDE                 ' Is this VB IDE?
            .SecurityLevel = eMan_Administrator   ' Appl authority required
            .CreateManifestFile                   ' Create/Update manifest file
        End With
        Set objManifest = Nothing

        InitComctl32       ' Initialize manifest file
        Unload frmSplash   ' Unload splash screen
        Load frmAbout      ' Load and hide
        Load frmMain       ' Load main form

    End If

Main_CleanUp:
    ' Verify class objects are freed from memory
    Set objOperSys = Nothing
    Set objManifest = Nothing
    Set mobjPrivileges = Nothing

    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

Main_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    TerminateProgram
    Resume Main_CleanUp

End Sub

' ***************************************************************************
' Routine:       TerminateProgram
'
' Description:   This routine will perform the shutdown process for this
'                application.  The proper sequence to follow is:
'
'                    1.  Deactivate and free from memory all global objects
'                        or classes
'                    2.  Verify there are no file handles left open
'                    3.  Deactivate and free from memory all form objects
'                    4.  Shut this application down
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub TerminateProgram()

    ' Update registry settings for this application
    frmMain.UpdateRegistry

    ' Free any global objects from memory.
    ' EXAMPLE:    Set gobjFSO = Nothing

    Close   ' Close all files opened by this application

    Set mobjPrivileges = New cPrivileges                  ' Instantiate class object
    mobjPrivileges.CheckPrivileges False, SE_DEBUG_NAME   ' Reset user privileges to what they were
    Set mobjPrivileges = Nothing                          ' Free class object from memory

    UnloadAllForms   ' Unload any forms from memory

    ' While in the VB IDE (VB Integrated Developement Environment),
    ' do not call ExitProcess API.  ExitProcess API will close all
    ' processes associated with this application including the IDE.
    ' No changes will be retained that were not previously saved.
    If mblnVB_IDE Then
        End    ' Terminate this application while in the VB IDE
    Else
        ' Close running application gracefully
        ExitProcess GetExitCodeProcess(GetCurrentProcess(), 0)
    End If

End Sub

' ***************************************************************************
' Routine:       UnloadAllForms
'
' Description:   Unload all active forms associated with this application.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub UnloadAllForms()

    ' Called by TerminateProgram()

    Dim frm As Form
    Dim ctl As Control

    On Error Resume Next

    ' Loop thru all active forms
    ' associated with this application
    For Each frm In Forms

        frm.Hide   ' Hide selected form

        ' Loop thru all active forms
        ' associated with this application
        For Each ctl In frm.Controls

            ' If control is timer
            If TypeOf ctl Is Timer Then
                ctl.Interval = 0      ' Set timer interval to 0
                ctl.Enabled = False   ' Disable timer
            End If

            Set ctl = Nothing   ' Free control from memory

        Next ctl

        Unload frm          ' Deactivate form object
        Set frm = Nothing   ' Free form object from memory
                            ' (prevents memory fragmentation)
    Next frm

    On Error GoTo 0   ' Nullify this error trap

End Sub

' ***************************************************************************
' Routine:       InitComctl32
'
' Description:   This will create the XP Manifest file and utilize it. You
'                will only see the results when the exe (not in the IDE)
'                is run.  This routine is usually called before any forms
'                are loaded.  (See modMain.bas)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jan-2006  Randy Birch
'              http://vbnet.mvps.org/
' 03-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 31-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated error trap by clearing error number
' ***************************************************************************
Public Sub InitComctl32()

    Dim typICC As INIT_COMMON_CTRLS

    On Error GoTo Use_Old_Version

    With typICC
        .dwSize = LenB(typICC)
        .dwICC = ICC_ALL_CLASSES
    End With

    ' VB will generate error 453 "Specified DLL function not found"
    ' if InitCommonControlsEx can't be located in the library.  An
    ' error is trapped and then original InitCommonControls is called
    ' instead below.
    If InitCommonControlsEx(typICC) = 0 Then
        InitCommonControls
    End If

    DoEvents
    On Error GoTo 0      ' Nullify this error trap
    Exit Sub

Use_Old_Version:
    Err.Clear            ' Clear any error codes
    InitCommonControls
    DoEvents
    On Error GoTo 0      ' Nullify this error trap

End Sub

' ***************************************************************************
' Routine:       FindRequiredFile
'
' Description:   Test to see if a required file is in the application folder
'                or in any of the folders in the PATH and WINDIR environment
'                variables.
'
'                If the required file is a registered file, access registry:
'
'                    Ref:  strFilename = "excel.exe"
'                         "Path" is subkey name with complete path w/o filename
'
'                    HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & strFilename, "Path"
'
' Syntax:        FindRequiredFile "msinfo32.exe", strPathFile
'
' Parameters:    strFilename - name of the file without path information
'                strFullPath - Optional - If found then the fully qualified
'                     path and filename are returned
'
' Returns:       TRUE  - Found the required file
'                FALSE - File could not be found
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 04-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 29-Dec-2013  Kenneth Ives  kenaso@tx.rr.com
'              Force a search of Windows and System32 folders if file is not
'              found.  Sometimes "PATH" variable becomes corrupted and
'              Windows is not in the "PATH".
' 08-May-2014  Kenneth Ives  kenaso@tx.rr.com
'              Add search of Windows, System32 and SysWOW64 folders
' 18-Sep-2014  Kenneth Ives  kenaso@tx.rr.com
'              Updated search logic
' ***************************************************************************
Public Function FindRequiredFile(ByVal strFileName As String, _
                        Optional ByRef strFullPath As String) As Boolean

    Dim blnFoundit       As Boolean      ' Flag (TRUE if found file else FALSE)
    Dim lngCount         As Long         ' String pointer position
    Dim lngIndex         As Long         ' Array index pointer
    Dim strPath          As String       ' Fully qualified search path
    Dim strMsgFmt        As String       ' Format each message line
    Dim strDosPath       As String       ' DOS environment variable
    Dim strSearched      As String       ' List of searched folders (will be displayed if not found)
    Dim strWindowsFolder As String       ' Windows folder
    Dim astrPath()       As String       ' List of folders to be searched

    On Error GoTo FindRequiredFile_Error

    strFullPath = ""   ' Verify empty variables
    strSearched = ""

    strMsgFmt = "!" & String$(70, "@")   ' Left justify data
    blnFoundit = False                   ' Preset flag to FALSE

    strPath = QualifyPath(App.Path)      ' Add trailing backslash to application folder

    ' Add current path, without file name, to list of searched folders
    strSearched = Format$(strPath, strMsgFmt) & vbNewLine

    strPath = strPath & strFileName      ' Append file name to path
    blnFoundit = IsPathValid(strPath)    ' Check selected folder

    ' Capture DOS environment variable 'PATH' and
    ' perform a search of the various folders
    If Not blnFoundit Then

        ' Capture DOS environment variable 'PATH' statement
        strDosPath = TrimStr(Environ$("PATH"))

        If Len(strDosPath) > 0 Then

            Erase astrPath()   ' Start with empty array
            lngCount = 0       ' Initialize array counter

            strDosPath = QualifyPath(strDosPath, ";")   ' Add trailing semi-colon
            astrPath() = Split(strDosPath, ";")         ' Load paths into array
            lngCount = UBound(astrPath)                 ' Number of entries in array

            For lngIndex = 0 To (lngCount - 1)

                strPath = astrPath(lngIndex)     ' Capture path
                strPath = GetLongName(strPath)   ' Format long path name

                ' Verify there is some data to work with
                If Len(strPath) > 0 Then

                    strPath = QualifyPath(strPath)   ' Add trailing backslash

                    ' Verify this folder exist
                    If IsPathValid(strPath) Then

                        ' Add current path, without file name, to list of searched folders
                        strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine

                        strPath = strPath & strFileName     ' Append file name to path
                        blnFoundit = IsPathValid(strPath)   ' Check selected folder

                        If blnFoundit Then
                            Exit For   ' Exit FOR..NEXT loop
                        End If
                    Else
                        ' Add current path, without file name, to list
                        ' of searched folders and update error message
                        strSearched = strSearched & Format$(strPath & " (Folder does not exist)", strMsgFmt) & vbNewLine
                    End If
                End If

            Next lngIndex
        Else
            ' Add current path, without file name, to list
            ' of searched folders and update error message
            strSearched = strSearched & Format$(Chr$(34) & "PATH" & Chr$(34) & _
                          " environment variable does not exists.", strMsgFmt) & vbNewLine
        End If
    End If

    ' Capture DOS environment variable 'WINDIR' and
    ' perform a search of Windows, System32 and SysWOW64
    ' folders.  Sometimes these folders are not in the
    ' 'PATH" variable and must be searched manually.
    If Not blnFoundit Then

        strPath = vbNullString           ' Verify empty variables
        strWindowsFolder = vbNullString

        ' Capture DOS environment variable "WINDIR" statement.
        ' This is the root folder for Windows.
        strWindowsFolder = TrimStr(Environ$("WINDIR"))

        ' See if "WINDIR" variable exist
        If Len(strWindowsFolder) > 0 Then

            ' Prepare to search Windows main folder (ex:  C:\Windows\)
            ' Format long folder name and add trailing backslash
            strWindowsFolder = QualifyPath(GetLongName(strWindowsFolder))

            ' Add current path, without file name, to list of searched folders
            strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine

            strPath = strWindowsFolder & strFileName   ' Append file name to path
            blnFoundit = IsPathValid(strPath)          ' Check Windows folder

            If Not blnFoundit Then

                ' Prepare to search System32 folder (ex:  C:\Windows\System32\)
                strPath = QualifyPath(strWindowsFolder & "System32")

                ' Add current path, without file name, to list of searched folders
                strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine

                strPath = strPath & strFileName     ' Append file name to path
                blnFoundit = IsPathValid(strPath)   ' Check System32 folder
            End If

            If Not blnFoundit Then

                ' Prepare to search SysWOW64 folder (ex:  C:\Windows\SysWOW64\)
                strPath = QualifyPath(strWindowsFolder & "SysWOW64")

                ' Add current path, without file name, to list of searched folders
                strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine

                strPath = strPath & strFileName     ' Append file name to path
                blnFoundit = IsPathValid(strPath)   ' Check SysWOW64 folder
            End If
        Else
            ' Add current path, without file name, to list
            ' of searched folders and update error message
            strSearched = strSearched & Format$(Chr$(34) & "WINDIR" & Chr$(34) & _
                          " environment variable does not exists.", strMsgFmt) & vbNewLine
        End If
    End If

FindRequiredFile_CleanUp:
    If blnFoundit Then
        strFullPath = strPath   ' Return full path/filename
    Else
        InfoMsg Format$("A required file that supports this application cannot be found.", strMsgFmt) & _
                vbNewLine & vbNewLine & _
                Format$(Chr$(34) & UCase$(strFileName) & Chr$(34) & _
                " not in any of these folders:", strMsgFmt) & vbNewLine & vbNewLine & _
                strSearched, , "File not found"
    End If

    FindRequiredFile = blnFoundit   ' Set status flag
    strSearched = vbNullString      ' Empty variable
    Erase astrPath()                ' Empty array

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

FindRequiredFile_Error:
    If Err.Number <> 0 Then
        Err.Clear
    End If

    Resume FindRequiredFile_CleanUp

End Function

' ***************************************************************************
' Procedure:     GetLongName
'
' Description:   The Dir() function can be used to return a long filename
'                but it does not include path information. By parsing a
'                given short path/filename into its constituent directories,
'                you can use the Dir() function to build a long path/filename.
'
' Example:       Syntax:
'                   GetLongName C:\DOCUME~1\KENASO\LOCALS~1\Temp\~ki6A.tmp
'
'                Returns:
'                   "C:\Documents and Settings\Kenaso\Local Settings\Temp\~ki6A.tmp"
'
' Parameters:    strShortName - Path or file name to be converted.
'
' Returns:       A readable path or file name.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jul-2004  http://support.microsoft.com/kb/154822
'              "How To Get a Long Filename from a Short Filename"
' 09-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 09-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added removal of all double quotes prior to formatting
' ***************************************************************************
Public Function GetLongName(ByVal strShortName As String) As String

    Dim strTemp     As String
    Dim strLongName As String
    Dim intPosition As Integer

    On Error Resume Next

    GetLongName = vbNullString
    strLongName = vbNullString

    ' Remove all double quotes
    strShortName = Replace(strShortName, Chr$(34), vbNullString)

    ' Add a backslash to short name, if needed,
    ' to prevent Instr() function from failing.
    strShortName = QualifyPath(strShortName)

    ' Start at position 4 so as to ignore
    ' "[Drive Letter]:\" characters.
    intPosition = InStr(4, strShortName, "\")

    ' Pull out each string between
    ' backslash character for conversion.
    Do While intPosition > 0

        strTemp = vbNullString   ' Init variable

        ' Progressively parse path to verify
        ' each portion does exist and
        ' capture its expanded version.
        strTemp = Dir$(Left$(strShortName, intPosition - 1), _
                       vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)

        ' If no data then exit this loop
        If Len(TrimStr(strTemp)) = 0 Then
            strShortName = vbNullString
            strLongName = vbNullString
            Exit Do   ' exit DO..LOOP
        End If

        ' Append new elongated portion to output string
        ' after converting it to propercase format.
        strLongName = strLongName & "\" & StrConv(strTemp, vbProperCase)

        ' Find next backslash
        intPosition = InStr(intPosition + 1, strShortName, "\")

    Loop

GetLongName_CleanUp:
    If Len(strShortName & strLongName) > 0 Then
        GetLongName = UCase$(Left$(strShortName, 2)) & strLongName
    Else
        GetLongName = "[Unknown]"
    End If

    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'                Shlwapi.dll first shipped with Internet Explorer 4.0
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))

End Function

' ***************************************************************************
' Routine:       AlreadyRunning
'
' Description:   This routine will determine if an application is already
'                active, whether it be hidden, minimized, or displayed.
'
' Parameters:    strTitle - partial/full name of application
'
' Returns:       TRUE  - Currently active
'                FALSE - Inactive
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-DEC-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function AlreadyRunning(ByVal strAppTitle As String) As Boolean

    Dim hMutex As Long

    Const ROUTINE_NAME As String = "AlreadyRunning"

    On Error GoTo AlreadyRunning_Error

    mblnVB_IDE = False  ' preset flags to FALSE
    AlreadyRunning = False

    ' Are we in VB development environment?
    mblnVB_IDE = IsVB_IDE

    ' Multiple instances can be run while
    ' in the VB IDE but not as an EXE
    If Not mblnVB_IDE Then

        ' Try to create a new Mutex handle
        hMutex = CreateMutex(ByVal 0&, 1, strAppTitle)

        ' Did mutex handle already exist?
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then

            ReleaseMutex hMutex     ' Release Mutex handle from memory
            CloseHandle hMutex      ' Close the Mutex handle
            Err.Clear               ' Clear any errors
            AlreadyRunning = True   ' prior version already active
        End If
    End If

AlreadyRunning_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

AlreadyRunning_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume AlreadyRunning_CleanUp

End Function

Private Function IsVB_IDE() As Boolean

    ' Called by AlreadyRunning()
    '
    ' 09-16-2000  Michael Culley  m_culley@one.net.au
    '             http://forums.devx.com/showthread.php?t=37676
    '
    ' Set DebugMode flag.  Call can only be successful if
    ' in the VB Integrated Development Environment (IDE).
    Debug.Assert SetTrue(IsVB_IDE) Or True

End Function

Private Function SetTrue(ByRef blnValue As Boolean) As Boolean

    ' Called by IsVB_IDE()
    '
    ' 09-16-2000  Michael Culley  m_culley@one.net.au
    '             http://forums.devx.com/showthread.php?t=37676
    '
    ' Can only be set to TRUE if Debug.Assert call is
    ' successful.  Call can only be successful if in
    ' the VB Integrated Development Environment (IDE).
    blnValue = True

End Function

' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function QualifyPath(ByVal strPath As String, _
                   Optional ByVal strChar As String = "\") As String

    strPath = TrimStr(strPath)

    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If

End Function

' ***************************************************************************
' Routine:       UnQualifyPath
'
' Description:   Removes a trailing character from the path
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to remove
'                          Default = "\"
'
' Returns:       Fully qualified path without a specific trailing character
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function UnQualifyPath(ByVal strPath As String, _
                     Optional ByVal strChar As String = "\") As String

    strPath = TrimStr(strPath)

    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        UnQualifyPath = Left$(strPath, Len(strPath) - 1)
    Else
        UnQualifyPath = strPath
    End If

End Function

' ***************************************************************************
' Routine:       SendEmail
'
' Description:   When the email hyperlink is clicked, this routine will fire.
'                It will create a new email message with the author's name in
'                the "To:" box and the name and version of the application
'                on the "Subject:" line.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Feb-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 25-Apr-2015  Kenneth Ives  kenaso@tx.rr.com
'              Added reference to GetDesktopWindow() API
' ***************************************************************************
Public Sub SendEmail()

    Dim strMail As String

    ' Create email heading for user
    strMail = "mailto:" & SUPPORT_EMAIL & "?subject=" & gstrVersion

    ' Call ShellExecute() API to create an email to the author
    ShellExecute GetDesktopWindow(), "open", strMail, _
                 vbNullString, vbNullString, vbNormalFocus

End Sub

' ***************************************************************************
' Routine:       ShrinkToFit
'
' Description:   This routine creates the ellipsed string by specifying
'                the size of the desired string in characters.  Adds
'                ellipses to a file path whose maximum length is specified
'                in characters.
'
' Parameters:    strPath - Path to be resized for display
'                intMaxLength - Maximum length of the return string
'
' Returns:       Resized path
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-May-2004  Randy Birch
'              http://vbnet.mvps.org/code/fileapi/pathcompactpathex.htm
' 22-Jun-2004  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function ShrinkToFit(ByVal strPath As String, _
                            ByVal intMaxLength As Integer) As String

    Dim strBuffer As String

    strPath = TrimStr(strPath)

    ' See if ellipses need to be inserted into the path
    If Len(strPath) <= intMaxLength Then
        ShrinkToFit = strPath
        Exit Function
    End If

    ' intMaxLength is the maximum number of characters to be contained in the
    ' new string, **including the terminating NULL character**. For example,
    ' if intMaxLength = 8, the resulting string would contain a maximum of
    ' seven characters plus the terminating null.
    '
    ' Because of this, add 1 to the value passed as intMaxLength to ensure
    ' the resulting string is the size requested.
    intMaxLength = intMaxLength + 1
    strBuffer = Space$(MAX_SIZE)
    PathCompactPathEx strBuffer, strPath, intMaxLength, 0&

    ' Return the readjusted data string
    ShrinkToFit = TrimStr(strBuffer)

End Function

' ***************************************************************************
' Routine:       DisableX
'
' Description:   Remove the "X" from the window and menu
'
'                A VB developer may find themselves developing an application
'                whose integrity is crucial, and therefore must prevent the
'                user from accidentally terminating the application during
'                its life, while still displaying the system menu.  And while
'                Visual Basic does provide two places to cancel an impending
'                close (QueryUnload and Unload form events) such a sensitive
'                application may need to totally prevent even activation of
'                the shutdown.
'
'                Although it is not possible to simply disable the Close button
'                while the Close system menu option is present, just a few
'                lines of API code will remove the system menu Close option
'                and in doing so permanently disable the titlebar close button.
'
' Parameters:    frmName - Name of form
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-Jul-1998  Randy Birch
'              RemoveMenu: Killing the Form's Close Menu and 'X' Button
'              http://vbnet.mvps.org/code/forms/killclose.htm
' ***************************************************************************
Public Sub DisableX(ByRef frmName As Form)

    Dim hMenu          As Long
    Dim lngMenuItemCnt As Long

    ' Obtain the handle to the form's system menu
    hMenu = GetSystemMenu(frmName.hWnd, 0&)

    If hMenu Then

        ' Obtain the handle to the form's system menu
        lngMenuItemCnt = GetMenuItemCount(hMenu)

        ' Remove the system menu Close menu item.
        ' The menu item is 0-based, so the last
        ' item on the menu is lngMenuItemCnt - 1
        RemoveMenu hMenu, lngMenuItemCnt - 1, _
                   MF_REMOVE Or MF_BYPOSITION

        ' Remove the system menu separator line
        RemoveMenu hMenu, lngMenuItemCnt - 2, _
                   MF_REMOVE Or MF_BYPOSITION

        ' Force a redraw of the menu. This
        ' refreshes the titlebar, dimming the X
        DrawMenuBar frmName.hWnd

    End If

End Sub

' ***************************************************************************
' Routine:       IsArrayInitialized
'
' Description:   This is an ArrPtr function that determines if the passed
'                array is initialized, and if so will return the pointer
'                to the safearray header. If the array is not initialized,
'                it will return zero. Normally you need to declare a VarPtr
'                alias into msvbvm50.dll or msvbvm60.dll depending on the
'                VB version, but this function will work with vb5 or vb6.
'                It is handy to test if the array is initialized as the
'                return value is non-zero.  Use CBool to convert the return
'                value into a boolean value.
'
'                This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
'                ex:  If CBool(IsArrayInitialized(array_being_tested)) Then ...
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Mar-2008  RD Edwards
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69970
' ***************************************************************************
Public Function IsArrayInitialized(ByVal avntData As Variant) As Long

    Dim intDataType As Integer   ' Variable must be a short integer

    On Error GoTo IsArrayInitialized_Exit

    IsArrayInitialized = 0  ' preset to FALSE

    ' Get the real VarType of the argument, this is similar
    ' to VarType(), but returns also the VT_BYREF bit
    CopyMemory intDataType, avntData, 2&

    ' if a valid array was passed
    If (intDataType And vbArray) = vbArray Then

        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena and Monte Hansen.
        CopyMemory IsArrayInitialized, ByVal VarPtr(avntData) + 8&, 4&

    End If

IsArrayInitialized_Exit:
    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       ShellAndWait
'
' Description:   Wait for a shelled application to close before continuing.
'
' Parameters:    strCmdLine - Data to be executed via the Shell process
'                lngDisplayMode- Optional - How to display shelled window
'                    Default = vbNormalFocus = 1
'                lngAttempts - Optional - Number of tries before forcing
'                    an exit from this routine.  Default = 3
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-Dec-2006  Randy Birch
'              http://vbnet.mvps.org/code/faq/getexitcprocess.htm
' 10-Sep-2014  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub ShellAndWait(ByVal strCmdLine As String, _
               Optional ByVal lngDisplayMode As Long = vbNormalFocus, _
               Optional ByVal lngAttempts As Long = 3)

    Dim lngExitCode  As Long
    Dim lngProcHwnd  As Long
    Dim lngProcessID As Long

    On Error GoTo ShellAndWait_CleanUp

    ' Start a shelled process and hide window
    lngProcessID = Shell(strCmdLine, lngDisplayMode)

    ' Capture shelled process handle
    lngProcHwnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, lngProcessID)

    ' The DoEvents statement relinquishes your application's control
    ' to allow Windows to process any pending messages or events for
    ' your application or any other running process.  Without this,
    ' your application will appear to lock up as the DO...LOOP
    ' essentially "grabs control" of the system.
    Do
        Call GetExitCodeProcess(lngProcHwnd, lngExitCode)
        DoEvents

        ' See if exit code designates FINISHED
        If lngExitCode <> STATUS_PENDING Then
            Exit Do   ' exit DO..LOOP
        Else
            Wait 1000                       ' Pause for one second
            lngAttempts = lngAttempts - 1   ' Decrement attempt counter
            DoEvents

            If lngAttempts < 1 Then
                Exit Do   ' exit DO..LOOP
            End If
        End If
    Loop

ShellAndWait_CleanUp:
    Call CloseHandle(lngProcHwnd)   ' Always release handle when not needed
    DoEvents
    On Error GoTo 0                 ' Nullify this error trap

End Sub

Public Sub Wait(ByVal lngMilliseconds As Long)

    Dim lngPause As Long

    ' Calculate a pause
    lngPause = GetTickCount() + lngMilliseconds

    Do
        DoEvents
    Loop While lngPause > GetTickCount()

End Sub


' ***************************************************************************
' Routine:       FindDefaultPath
'
' Description:   Find default path for a specific program by accessing
'                the registry.  For example, the default browser.  When
'                using API ShellExecute() command, the browser will
'                sometimes hang and show a blank page if it is not
'                already active.
'
' Parameters:    lngHiveID  - Numerical locaton of main hive in registry
'                strSection - Section path to be queried
'                strkeyName - Optional - Specific key value name to be
'                             queried.  Default = Null string
'
' Returns:       If successful, default path of queried item else
'                return an empty string.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-Sep-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function FindDefaultPath(ByVal lngHiveID As Long, _
                                ByVal strSection As String, _
                       Optional ByVal strKeyName As String = "") As String

    Dim lngStart     As Long
    Dim lngPathLen   As Long
    Dim lngPointer   As Long
    Dim lngDataType  As Long
    Dim lngKeyHandle As Long
    Dim strPath      As String

    FindDefaultPath = ""         ' Verify return is empty
    strPath = Space$(MAX_SIZE)   ' Preload with spaces
    lngPathLen = MAX_SIZE        ' Preset length of data

    Select Case UCase$(strKeyName)
           Case "@", "*", "(DEFAULT)"
                strKeyName = ""
    End Select

    ' Open section and capture key handle
    If RegOpenKeyEx(lngHiveID, strSection, 0&, KEY_READ, lngKeyHandle) = ERROR_SUCCESS Then

        ' Capture data type
        If RegQueryValueEx(lngKeyHandle, strKeyName, ByVal 0&, _
                           lngDataType, ByVal strPath, lngPathLen) = ERROR_SUCCESS Then

            ' Verify data type
            If (lngDataType = REG_SZ) Or _
               (lngDataType = REG_EXPAND_SZ) Then

                strPath = Space$(MAX_SIZE)   ' Preload with spaces
                lngPathLen = MAX_SIZE        ' Preset length of data

                ' Capture path and data length
                If RegQueryValueEx(lngKeyHandle, strKeyName, ByVal 0&, _
                                   lngDataType, ByVal strPath, lngPathLen) = ERROR_SUCCESS Then

                    strPath = TrimStr(Left$(strPath, lngPathLen))   ' Remove unwanted characters

                    If Len(strPath) = 0 Then
                        strPath = ""   ' No data found
                    Else
                        ' Look for first double quote
                        lngPointer = InStr(1, strPath, Chr$(34), vbBinaryCompare)

                        If lngPointer > 0 Then
                            lngStart = lngPointer + 1    ' Update start position

                            ' Look for next double quote
                            lngPointer = InStr(lngStart, strPath, Chr$(34), vbBinaryCompare)

                            strPath = Mid$(strPath, lngStart - 1, lngPointer)
                        End If
                    End If
                Else
                    strPath = ""   ' Bad RegQueryValueEx() read
                End If
            Else
                strPath = ""   ' Wrong data type
            End If
        Else
            strPath = ""   ' Bad RegQueryValueEx() read
        End If
    Else
        strPath = ""   ' Invalid section path
    End If

FindDefaultPath_CleanUp:
    If lngKeyHandle > 0 Then
        RegCloseKey lngKeyHandle   ' Always close key when not needed
        lngKeyHandle = 0
    End If

    FindDefaultPath = strPath  ' Return path

End Function
