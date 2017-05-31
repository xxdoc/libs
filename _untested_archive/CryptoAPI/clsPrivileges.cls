VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrivileges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        cPrivileges (clsPrivileges.cls)
'
' Description:   Perform temporary update of user privileges to perform
'                Reading and writing in resticted areas.  Since Windows 8,
'                C:\Program Files\ and C:\Program Files (x96)\ have become
'                restrictive areas ecept for the initail install or removal.
'                This includes creating a manifest file.  Now if I restrict
'                this application to Windows Vista or newer, then I can
'                install a manifest file at the same time and this module
'                would not be needed.
'
'                Verify user has Administrator authority.  If so, set this
'                application to always run as Administrator.
'
' References:    Privilege Constants
'                https://msdn.microsoft.com/en-us/library/windows/desktop/bb530716(v=vs.85).aspx
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Jun-2016  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 18-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated IsAdmin() routine concerning registry check for
'              proper key value
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME                   As String = "clsPrivileges"

  ' Used for setting user privileges
  Private Const SE_PRIVILEGE_ENABLED          As Long = &H2&
  Private Const TOKEN_QUERY                   As Long = &H8&
  Private Const TOKEN_ADJUST_PRIVILEGES       As Long = &H20&

  ' Registry Hive identifiers
  Private Const HKEY_CURRENT_USER             As Long = &H80000001   ' HKCU
  Private Const HKEY_LOCAL_MACHINE            As Long = &H80000002   ' HKLM
  Private Const HKEY_USERS                    As Long = &H80000003   ' HKUS

  ' Used to report API errors
  Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
  Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&

' ***************************************************************************
' Type structures
' ***************************************************************************
  Private Type LUID
      LowPart  As Long
      HighPart As Long
  End Type

  Private Type LUID_AND_ATTRIBUTES
      pLuid      As LUID
      Attributes As Long
  End Type

  Private Type TOKEN_PRIVILEGES
      PrivilegeCount     As Long
      Privileges(0 To 0) As LUID_AND_ATTRIBUTES
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The GetCurrentProcess function returns a pseudohandle for the current
  ' process. A pseudohandle is a special constant that is interpreted as
  ' the current process handle. The calling process can use this handle to
  ' specify its own process whenever a process handle is required. The
  ' pseudohandle need not be closed when it is no longer needed.
  Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

  ' OpenProcessToken function opens the access token associated with
  ' a process. If the function succeeds, the return value is nonzero.
  Private Declare Function OpenProcessToken Lib "advapi32.dll" _
          (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
          TokenHandle As Long) As Long

  ' LookupPrivilegeValue function retrieves the locally unique identifier
  ' (LUID) used on a specified system to locally represent the specified
  ' privilege name. If the function succeeds, the function returns nonzero.
  Private Declare Function LookupPrivilegeValue Lib "advapi32" _
          Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
          ByVal lpName As String, lpLuid As LUID) As Long

  ' AdjustTokenPrivileges function enables or disables privileges in the
  ' specified access token. Enabling or disabling privileges in an access
  ' token requires TOKEN_ADJUST_PRIVILEGES access. If the function succeeds,
  ' the return value is nonzero.
  Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
          (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
          NewState As Any, ByVal BufferLength As Long, _
          PreviousState As Any, ByVal ReturnLength As Long) As Long

  ' The FormatMessage function formats a message string. The function requires
  ' a message definition as input. The message definition can come from a
  ' buffer passed into the function. It can come from a message table resource
  ' in an already-loaded module. Or the caller can ask the function to search
  ' the system's message table resource(s) for the message definition. The
  ' function finds the message definition in a message table resource based on
  ' a message identifier and a language identifier. The function copies the
  ' formatted message text to an output buffer, processing any embedded insert
  ' sequences if requested.
  Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
          (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
          ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
          Arguments As Long) As Long

  ' CloseHandle invalidates the specified object handle, decrements
  ' the object’s handle count, and performs object retention checks.
  ' Once the last handle to an object is closed, the object is removed
  ' from the operating system. In other words, closes an open file.
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Boolean)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m bln IsAdmin
' Variable name:     mblnIsAdmin
' ***************************************************************************
  Private mblnVB_IDE     As Boolean            ' Operating in VB Developement?
  Private mblnIsAdmin    As Boolean            ' Return TRUE/FALSE
  Private mblnWindows7   As Boolean            ' Test for Windows 7
  Private mblnWindows8   As Boolean            ' Test for Windows 8
  Private mblnWindows10  As Boolean            ' Test for Windows 10 or Beta version
  Private mblnRegUpdated As Boolean            ' Test if registry was updated
  Private mstrComputer   As String             ' Name of this computer
  Private mstrUserName   As String             ' Name of current logon user
  Private mtypOldPriv    As TOKEN_PRIVILEGES   ' Backup copy of original privileges


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

' Data for these properties passed from modMain.Main() routine

Public Property Let bWindows10(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.bWindows10 = True
    mblnWindows10 = blnData
End Property

Public Property Let bWindows8(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.bWindows8 = True
    mblnWindows8 = blnData
End Property

Public Property Let bWindows7(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.bWindows7 = True
    mblnWindows7 = blnData
End Property

Public Property Let bVB_IDE(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.        Syntax: X.bVB_IDE = True
    mblnVB_IDE = blnData
End Property

Public Property Get bIsAdministrator() As Boolean
    ' used when a call to this property to find its value.
    ' Syntax:  calling_variable = X.bIsAdministrator
    bIsAdministrator = mblnIsAdmin
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CheckPrivileges
'
' Description:   Possibly change user privileges and if requested, see if
'                user is a member of the Administrators Group then update
'                this application's authority.
'
' Parameters:    blnEnablePrivilege - Toggle privilege on/off (TRUE/FALSE)
'                strPrivilege - Privilege name to be temporarily changed
'                blnTestForAdmin - Optional - Test to see if user has Admin
'                    authority to perform updates in restricted areas
'                    (Needed for Windows 8 and newer to perform writing
'                    to areas like C:\Program Files (x86)\ folders)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Jun-2016  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub CheckPrivileges(ByVal blnEnablePrivilege As Boolean, _
                           ByRef strPrivilege As String, _
                  Optional ByRef blnTestForAdmin As Boolean = False)

    Const ROUTINE_NAME As String = "CheckPrivileges"

    ' Temporarily update user privileges
    If AdjustPrivileges(blnEnablePrivilege, strPrivilege) Then

        ' Do we need to test for a member of Administrator Group?
        If blnTestForAdmin Then

            mblnIsAdmin = IsAdmin()

            Select Case mblnIsAdmin

                   Case True
                        If mblnVB_IDE Then
                            Exit Sub
                        Else
                            If mblnRegUpdated Then
                                Exit Sub
                            Else
                                InfoMsg "This application may not function correctly possibly" & vbNewLine & _
                                        "due to registry not being properly updated." & _
                                        vbNewLine & vbNewLine & _
                                        "Recommended action:" & vbNewLine & _
                                        vbTab & "1.  Close this application" & vbNewLine & _
                                        vbTab & "2.  Right mouse click executable name in" & vbNewLine & _
                                        vbTab & Space$(5) & "Windows Explorer" & vbNewLine & _
                                        vbTab & "3.  Select option to " & Chr$(34) & "Run as Administrator" & Chr$(34) & _
                                        vbNewLine & vbNewLine & vbNewLine & _
                                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , "Advisory message"
                            End If
                        End If

                   Case False   ' Due to UAC restrictions when using Windows 8, 8.1, or 10,
                                ' it must be determined if user has Admin authority to
                                ' write\update data in a folder within a restsricted area
                                ' such as "C:\Program Files (x86)\appl_name".
                        InfoMsg "This application may not function correctly possibly" & vbNewLine & _
                                "due to current logged in user " & vbNewLine & _
                                vbTab & Chr$(34) & UCase$(mstrUserName) & Chr$(34) & vbNewLine & _
                                "may not have sufficient rights." & vbNewLine & vbNewLine & _
                                "If running Windows 8 or newer then Microsoft has split" & vbNewLine & _
                                "Administrator authority via User Access Control (UAC)." & vbNewLine & vbNewLine & _
                                "Recommended action:" & vbNewLine & _
                                vbTab & "1.  Close this application" & vbNewLine & _
                                vbTab & "2.  Right mouse click executable name in" & vbNewLine & _
                                vbTab & Space$(5) & "Windows Explorer" & vbNewLine & _
                                vbTab & "3.  Select option to " & Chr$(34) & "Run as Administrator" & Chr$(34) & _
                                vbNewLine & vbNewLine & vbNewLine & _
                                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , "Advisory message"
            End Select
        End If
    End If

End Sub


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:       AdjustPrivileges
'
' Description:   The AdjustPrivileges function takes a single Boolean
'                parameter which indicates whether to toggle this privilege
'                on or off.  The meat of this routine lies in the
'                AdjustTokenPrivileges calls.
'
'                AdjustTokenPrivileges API requires you pass it a locally
'                unique identifier (LUID) for the privilege you would like
'                to obtain or release. These privileges have all been
'                assigned constant string values, which are used as look-up
'                references.
'
' Parameters:    blnEnablePrivilege - True - Set privilege level
'                            False - Restore to previous level
'                strPrivilege - Privilege name to be adjust to.
'
' Returns:       True if successful else False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 27-Jul-2010  Karl E. Peterson
'              Creating Admin Tools in a Least Privileged World
'              https://visualstudiomagazine.com/articles/2010/07/27/create-admin-tools.aspx
' 02-Oct-2016  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function AdjustPrivileges(ByVal blnEnablePrivilege As Boolean, _
                                  ByRef strPrivilege As String) As Boolean

    ' Called by CheckPrivileges()

    Dim lngHandle        As Long
    Dim lngDesiredAccess As Long
    Dim typLUID          As LUID
    Dim typNewPriv       As TOKEN_PRIVILEGES

    Const ROUTINE_NAME   As String = "AdjustPrivileges"

    On Error GoTo AdjustPrivileges_CleanUp

    AdjustPrivileges = False   ' Preset to FALSE

    If Len(TrimStr(strPrivilege)) = 0 Then

        InfoMsg "Cannot identify type of privilege to use." & _
                 vbNewLine & vbNewLine & _
                 "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo AdjustPrivileges_CleanUp
    End If

    ' Create handle to query and adjust process tokens
    lngDesiredAccess = TOKEN_QUERY Or TOKEN_ADJUST_PRIVILEGES

    ' Capture process handle for this application
    If OpenProcessToken(GetCurrentProcess(), lngDesiredAccess, lngHandle) Then

        ' Get LUID for higher privilege token
        If LookupPrivilegeValue(vbNullString, strPrivilege, typLUID) Then

            ' Elevate user privilege
            If blnEnablePrivilege Then

                ' Prepare new privilege
                With typNewPriv
                    .PrivilegeCount = 1

                    With .Privileges(0)
                        .Attributes = SE_PRIVILEGE_ENABLED
                        .pLuid.HighPart = typLUID.HighPart
                        .pLuid.LowPart = typLUID.LowPart
                    End With
                End With

                ' Enable new priviledge (Successful call is nonzero)
                If AdjustTokenPrivileges(lngHandle, False, _
                                         typNewPriv, Len(typNewPriv), _
                                         ByVal 0&, ByVal 0&) <> 0 Then

                    AdjustPrivileges = True
                Else
                    InfoMsg "Failed to elevate user privileges.  This" & vbNewLine & _
                            "application may not function as desired." & vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                End If

            Else
                ' Restore prior privilege settings (Successful call is nonzero)
                If AdjustTokenPrivileges(lngHandle, False, _
                                         mtypOldPriv, Len(mtypOldPriv), _
                                         ByVal 0&, ByVal 0&) <> 0 Then

                    AdjustPrivileges = True
                End If

            End If   ' blnEnablePrivilege
        Else
            InfoMsg "Cannot determine current process. " & vbNewLine & _
                    "[" & WinError(Err.LastDllError) & "]" & _
                     vbNewLine & vbNewLine & _
                     "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

        End If   ' LookupPrivilegeValue

    Else
        InfoMsg "Cannot find " & Chr$(34) & strPrivilege & Chr$(34) & " privilege value. " & _
                vbNewLine & "[" & WinError(Err.LastDllError) & "]" & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

    End If   ' OpenProcessToken

AdjustPrivileges_CleanUp:
    If lngHandle <> 0 Then
        Call CloseHandle(lngHandle)   ' Close process handle
        lngHandle = 0                 ' Reset to zero
    End If

    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       IsAdmin
'
' Description:   This routine returns TRUE if the caller's process is a
'                member of the Administrators local group.
'
'                With Windows 8 and newer, Microsoft will split Admin
'                users rights via User Access Control (UAC).  Recommendation
'                is the user will have to right-mouse click the application
'                executable and either
'
'                    Select the option to 'Run as Administrator'
'                            or
'                    Select 'Properties>>Compatibility' tab
'                    Check the box to 'Run this program as an administrator'
'
' Returns:       TRUE if member of Administrator Group else FALSE
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-Oct-2015  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 18-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated registry check for proper key value
' ***************************************************************************
Private Function IsAdmin() As Boolean

    ' Called by CheckPrivileges()

    Dim strValue    As String
    Dim strGrpName  As String
    Dim strValName  As String
    Dim strKeyPath1 As String
    Dim strKeyPath2 As String
    Dim objWMI      As Object
    Dim objReg      As Object
    Dim objMember   As Object
    Dim objMembers  As Object
    Dim objAdminGrp As Object

    On Error Resume Next

    Set objWMI = Nothing      ' Verify objects are freed from memory
    Set objReg = Nothing
    Set objMember = Nothing
    Set objMembers = Nothing
    Set objAdminGrp = Nothing

    IsAdmin = False          ' Preset to FALSE
    mblnRegUpdated = False

    ' This code does not care what the country locale
    ' uses for the name of their Administrators group
    Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & mstrComputer & "/root/cimv2")

    Set objMember = objWMI.Get("Win32_SID.SID='S-1-5-32-544'")   ' SID for Administrators Group ('S-1-5-32-544')
    strGrpName = objMember.AccountName                           ' Local name of Administrators group

    Set objMember = Nothing   ' Free objects from memory
    Set objWMI = Nothing

    Set objAdminGrp = GetObject("WinNT://" & mstrComputer & "/" & strGrpName)
    Set objMembers = objAdminGrp.Members

    ' Find current user name
    For Each objMember In objMembers

        If StrComp(UCase$(objMember.Name), UCase$(mstrUserName), vbBinaryCompare) = 0 Then

            IsAdmin = True   ' Found user name

            ' If this application is being run within the VB IDE
            ' (Integrated Developement Environment) then no need
            ' to update the registry
            If mblnVB_IDE Then
                Exit For   ' Exit For..Next loop
            Else
                strKeyPath1 = "S-1-5-21-3569730504-1347434104-3088361220-1001\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
                strKeyPath2 = "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
                strValName = QualifyPath(App.Path) & App.EXEName & ".exe"
                strValue = vbNullString

                Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & mstrComputer & "/root/default:StdRegProv")

                With objReg

                    ' See if key already exist
                    .GetStringValue HKEY_LOCAL_MACHINE, strKeyPath2, strValName, strValue

                    strValue = TrimStr(strValue)   ' Remove any unwanted leading/trailing chars

                    ' Data found in key
                    If Len(strValue) > 0 then

                        ' Updated with proper value?
                        If InStr(1, UCase$(strValue), "RUNASADMIN", vbBinaryCompare) > 0 Then
                            mblnRegUpdated = True   ' Registry previously updated
                            Exit For                ' Exit For..Next loop
                        End If
                    End If

                    ' Invalid or no key value found
                    DoEvents
                    strValue = vbNullString   ' Verify empty variable

                    ' Use appropriate value based
                    ' on operating system version
                    If mblnWindows10 Then
                        strValue = "@ RUNASADMIN"
                    ElseIf mblnWindows8 Then
                        strValue = "~ RUNASADMIN"
                    ElseIf mblnWindows7 Then
                        strValue = "# RUNASADMIN"
                    Else
                        strValue = "RUNASADMIN"
                    End If

                    ' This is the correct sequence of events.  To verify,
                    ' close this application after it has fully loaded.
                    ' Right mouse click the EXE file and select
                    ' Properties>>Compatibility tab.  There should be a
                    ' checkmark next to "Run this application as an
                    ' Administrator".
                    If .CreateKey(HKEY_USERS, strKeyPath1) = 0 Then
                        .SetStringValue HKEY_USERS, strKeyPath1, strValName, strValue
                    End If

                    If .CreateKey(HKEY_CURRENT_USER, strKeyPath2) = 0 Then
                        .SetStringValue HKEY_CURRENT_USER, strKeyPath2, strValName, strValue
                    End If

                    If .CreateKey(HKEY_LOCAL_MACHINE, strKeyPath2) = 0 Then
                        .SetStringValue HKEY_LOCAL_MACHINE, strKeyPath2, strValName, strValue
                    End If

                    ' Verify key was updated
                    DoEvents
                    strValue = vbNullString
                    .GetStringValue HKEY_LOCAL_MACHINE, strKeyPath2, strValName, strValue

                    strValue = TrimStr(strValue)   ' Remove any unwanted leading/trailing chars

                    ' Updated with proper value?
                    If InStr(1, UCase$(strValue), "RUNASADMIN", vbBinaryCompare) > 0 Then

                        mblnRegUpdated = True      ' Registry has been updated
                        Set objWMI = Nothing       ' Free objects from memory
                        Set objReg = Nothing
                        Set objMember = Nothing
                        Set objMembers = Nothing
                        Set objAdminGrp = Nothing

                        InfoMsg "Registry has been updated with proper" & vbNewLine & _
                                "authority to execute this application." & vbNewLine & vbNewLine & _
                                "Please close and restart this application.", , , 4
                        TerminateProgram   ' Close this application

                    Else
                        mblnRegUpdated = False   ' Registry has NOT been updated
                    End If

                End With

                Exit For   ' Exit For..Next loop

            End If
        End If

    Next objMember

IsAdmin_CleanUp:
    Set objWMI = Nothing      ' Free objects from memory
    Set objReg = Nothing
    Set objMember = Nothing
    Set objMembers = Nothing
    Set objAdminGrp = Nothing

    On Error GoTo 0   ' Nullify this error trap

End Function

Private Function WinError(ByVal lngLastDLLError As Long) As String

    ' Called by AdjustPrivileges()

    Dim strBuffer As String
    Dim lngLength As Long

    ' Return error message associated with LastDLLError
    strBuffer = String$(256, 0)
    lngLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                              FORMAT_MESSAGE_IGNORE_INSERTS, _
                              0, lngLastDLLError, 0&, _
                              strBuffer, Len(strBuffer), ByVal 0)
    If lngLength Then
        WinError = Left$(strBuffer, lngLength)
    End If

End Function

Private Sub CaptureComputerName()

    ' Called by Class_Initialize()

    Dim objWMI As Object

    Set objWMI = CreateObject("Wscript.Network")

    With objWMI
        mstrComputer = .ComputerName   ' Capture current computer name
        mstrUserName = .UserName       ' Name of current logged in user
    End With

    Set objWMI = Nothing   ' Free object from memory

    If Len(TrimStr(mstrComputer)) = 0 Then
        mstrComputer = "."
    End If

    If Len(TrimStr(mstrUserName)) = 0 Then
        mstrUserName = "UNKNOWN"
    End If

End Sub

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    mblnIsAdmin = False   ' Preset property values
    bWindows7 = False
    bWindows8 = False
    bWindows10 = False

    CaptureComputerName   ' Capture name of logged in user and computer

End Sub
