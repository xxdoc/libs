VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cManifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***************************************************************************
' Module:       cManifest  (clsManifest.cls)
'
' Important:     If you should get a Side-by-side error message, after
'                making changes to this module, try this:
'
'                1.  Delete current manifest file (ex: app_name.exe.manifest)
'                2.  Restart this application.  A new manifest file should be
'                    created.
'
' Description:  Names of elements and attributes in a manifest file are case
'               sensitive. The values of elements and attributes are not case
'               sensitive, except for the value of the type attribute.
'
'               A manifest is a XML file that contains settings that informs
'               Windows how to handle a program when it is started. The
'               manifest can be embedded inside the program file (as a resource)
'               or it can be located in a separate external XML file. If the
'               manifest is placed in a separate file, then the file must be
'               located in the same folder as the executable file and it must
'               have same filename as the program file, but with a ".manifest"
'               filename extension added at the end (e.g "MYAPP.EXE.manifest").
'
'               Manifests are also often used to inform Windows which version
'               of a DLL a program is dependent of. By using manifests same
'               DLL can exists in the computer in different versions and the
'               program chooses which version of the DLL to load. This is
'               sometimes referred to as the side-by-side technology in Windows.
'               DLL's that exists in different versions in the computer are
'               sometimes referred to as side-by-side assemblies.
'
'               One common use of manifests and the side-by-side technology
'               is to inform Windows which version of the Windows common
'               controls the program need to use. The Windows common controls
'               are used to draw menus, dialog boxes, buttons, etc. in an
'               application and two different styles can be used: a classic
'               visual style and a modern visual style. The "modern visual
'               style" is the modern 3D style that was introduced with
'               Windows XP and that has evolved even more in Windows Vista
'               or Windows 7, with shadows, light effects and metal effects.
'
'               These files have been around for a while, and the new
'               section added for UAC control is the requestedExecutionLevel
'               element. The most effected by UAC Control are Windows Vista
'               and newer versions. The level attribute may be one of the
'               following values:
'
'               --------------------------------------------------------------
'               Level               Description
'               --------------------------------------------------------------
'               asInvoker             Does not require elevation, runs without
'                                     requesting elevation using privileges of
'                                     its parent process.
'
'               highestAvailable      Requests the highest available privilege
'                                     tokens of its parent process. An admin
'                                     account will attempt to elevate to full
'                                     administrator level, but a standard
'                                     account will only request elevation to
'                                     its own highest set of access tokens.
'
'               requireAdministrator  Requires elevation to full administrator
'                                     privileges.
'               --------------------------------------------------------------
'
'               Most applications should use the asInvoker level, as this
'               will not prompt the user for elevation every time it is run,
'               and will use the standard set of privileges to access the
'               unprotected resources it needs. Only applications that require
'               access to protected resources should use the higher access levels.
'
'               The uiAccess attribute determines whether the application
'               requires access to any protected UI elements, such as system
'               dialog boxes or higher-level processes. Only signed applications
'               may do this as it requires additional security tokens. This value
'               defaults to false.
'
' References:   Application Manifests
'               http://msdn.microsoft.com/en-us/library/aa374191(v=vs.85).aspx
'
'               What is a Manifest (in Windows)?
'               http://www.samlogic.net/articles/manifest.htm
'
'               Making Your Application UAC Aware
'               TheCodeKing,  10 Mar 2007
'               http://www.codeproject.com/Articles/17968/Making-Your-Application-UAC-Aware
'
'               Targeting your application for Windows 8.1
'               http://msdn.microsoft.com/en-us/library/windows/desktop/dn481241(v=vs.85).aspx
'
'               SupportedOS manifest for Windows 10
'               https://social.msdn.microsoft.com/Forums/en-US/07cbfc3a-bced-45b7-80d2-a9d32a7c95d4/supportedos-manifest-for-windows-10?forum=windowsgeneraldevelopmentissues
'
'               Manifest Madness
'               This blog post discusses Application Manifest elements for
'               use in Win32 desktop applications. This is not about the App
'               package manifest (aka AppX) used for Windows Store apps for
'               Windows 8.x.
'               http://blogs.msdn.com/b/chuckw/archive/2013/09/10/manifest-madness.aspx
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added IsPathValid() routine.
' 28-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic in NeedManifest() routine
' 31-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated error trap by clearing error number in InitComctl32()
'              routine.
' 28-Oct-2013  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote data for inside of manifest file
' 28-Jan-2014  Kenneth Ives  kenaso@tx.rr.com
'              Updated data for inside of manifest file when application
'              began recieving side-by-side error message because of
'              previous manifest file being present.
' 13-Oct-2014  Kenneth Ives  kenaso@tx.rr.com
'              Evaluate operating system to determine if a manifest file
'              needs to be created
' 12-Jul-2015  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote CreateManifestFile() routine
' 12-Jun-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added properties to obtain current operating system
'              information from modMain.bas
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Type structures
' ***************************************************************************
  Private Type VS_FIXEDFILEINFO
      dwSignature        As Long
      dwStrucVersion     As Long
      dwFileVersionMS    As Long
      dwFileVersionLS    As Long
      dwProductVersionMS As Long
      dwProductVersionLS As Long
      dwFileFlagsMask    As Long
      dwFileFlags        As Long
      dwFileOS           As Long
      dwFileType         As Long
      dwFileSubtype      As Long
      dwFileDateMS       As Long
      dwFileDateLS       As Long
  End Type

' ***************************************************************************
' Enumberations
' ***************************************************************************
  Public Enum enumSecurityLevel
      eMan_Invoker         ' 0
      eMan_HighestAvail    ' 1
      eMan_Administrator   ' 2
  End Enum

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' GetFileVersionInfoSize function determines whether the operating system
  ' can retrieve version information for a specified file. If version
  ' information is available, GetFileVersionInfoSize returns the size, in
  ' bytes, of that information.
  Private Declare Function GetFileVersionInfoSize Lib "Version.dll" _
          Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
          lpdwHandle As Long) As Long

  ' GetFileVersionInfo function retrieves version information for the
  ' specified file.
  Private Declare Function GetFileVersionInfo Lib "Version.dll" _
          Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, _
          ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long

  ' VerQueryValue function retrieves specified version information from
  ' the specified version-information resource. To retrieve the appropriate
  ' resource, before you call VerQueryValue, you must first call the
  ' GetFileVersionInfoSize function, and then GetFileVersionInfo function.
  Private Declare Function VerQueryValue Lib "Version.dll" _
          Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
          lplpBuffer As Any, puLen As Long) As Long

  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
          (Destination As Any, Source As Any, ByVal Length As Long)

  ' ZeroMemory function fills a block of memory with zeros.
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Long)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m lng MajorVer
' Variable name:     mlngMajorVer
' ***************************************************************************
  Private mlngMajorVer      As Long
  Private mlngMinorVer      As Long
  Private mstrSecurityLevel As String
  Private mblnVB_IDE        As Boolean


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

' Data passed from modMain.bas | Sub Main()

Public Property Let MajorVersion(ByVal lngData As Long)
    mlngMajorVer = lngData
End Property

Public Property Let MinorVersion(ByVal lngData As Long)
    mlngMinorVer = lngData
End Property

Public Property Let bVB_IDE(ByVal blnData As Boolean)
    mblnVB_IDE = blnData
End Property

Public Property Let SecurityLevel(ByVal lngData As Long)

    Select Case lngData

           ' Does not require elevation, runs without requesting
           ' elevation using privileges of its parent process
           Case eMan_Invoker: mstrSecurityLevel = "asInvoker"

           ' Requests highest available privilege tokens of its
           ' parent process.  An admin account will attempt to
           ' elevate to full administrator level, but a standard
           ' account will only request elevation to its own highest
           ' set of access tokens.
           Case eMan_HighestAvail: mstrSecurityLevel = "highestAvailable"

           ' Requires elevation to full administrator privileges
           Case eMan_Administrator: mstrSecurityLevel = "requireAdministrator"

           Case Else: mstrSecurityLevel = "asInvoker"   ' Default value
    End Select

End Property


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CreateManifestFile
'
' Description:   Create a manifest file based on the Windows major version
'                number.
'
' Important:     If you should get a Side-by-side error message, after
'                making changes to this module, try this:
'
'                1.  Delete current manifest file (ex: app_name.exe.manifest)
'                2.  Restart this application.  A new manifest file should be
'                    created.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jan-2006  Randy Birch
'              http://vbnet.mvps.org/
' 28-Oct-2013  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 13-Oct-2014  Kenneth Ives  kenaso@tx.rr.com
'              Evaluate operating system to determine if a manifest file
'              needs to be created
' 18-Aug-2015  Kenneth Ives  kenaso@tx.rr.com
'              Updated to test for Windows major version six or newer
' 18-Jun-2016  Kenneth Ives  kenaso@tx.rr.com
'              Determine which security level to invoke based on operating
'              system
'              Note:  processorArchitecture="*"  Will run on either 32 or 64
'                     bit architecture
' ***************************************************************************
Public Function CreateManifestFile() As Boolean

    Dim hFile            As Long
    Dim strXML           As String
    Dim strExeName       As String
    Dim strFileName      As String
    Dim strAppVersion    As String
    Dim strDescription   As String
    Dim blnVer6orNewer   As Boolean
    Dim blnWin7orEarlier As Boolean

    On Error GoTo CreateManifestFile_CleanUp

    CreateManifestFile = False                       ' Preset to FALSE
    strXML = vbNullString                            ' Verify output string is empty
    strExeName = App.EXEName                         ' EXE name without extension
    strAppVersion = FileVersion(strExeName)          ' Get appl EXE major version number
    strDescription = "Kens.Software." & strExeName   ' Enter unique name here
    strFileName = QualifyPath(App.Path) & strExeName & ".exe.manifest"

    ' See if a manifest file is needed
    If NeedManifest(blnVer6orNewer, blnWin7orEarlier) Then

        ' Create manifest file for Vista, 7, 8, 8.1, 10
        ' Manifest file contents are mostly case sensitive
        If blnVer6orNewer Then

            ' Contents of manifest file if Windows major version is 6
            strXML = strXML & "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & _
                              "UTF-8" & Chr$(34) & " standalone=" & Chr$(34) & "yes" & Chr$(34) & "?>" & vbNewLine

            If blnWin7orEarlier Then
                ' Vista, Win 7
                strXML = strXML & "<assembly xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v1" & Chr$(34) & _
                                  " manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & ">" & vbNewLine
            Else
                ' Win 8, 8.1, 10
                strXML = strXML & "<assembly manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & _
                                  " xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v1" & Chr$(34) & _
                                  " xmlns:asmv3=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v3" & Chr$(34) & ">" & vbNewLine
            End If

            ' Minimum version of this application
            strXML = strXML & "  <assemblyIdentity" & vbNewLine
            strXML = strXML & "    type=" & Chr$(34) & "win32" & Chr$(34) & vbNewLine
            strXML = strXML & "    name=" & Chr$(34) & strExeName & Chr$(34) & vbNewLine
            strXML = strXML & "    version=" & Chr$(34) & strAppVersion & Chr$(34) & vbNewLine
            strXML = strXML & "    processorArchitecture=" & Chr$(34) & "*" & Chr$(34) & vbNewLine
            strXML = strXML & "  />" & vbNewLine

            ' Company and application name
            strXML = strXML & "  <description>" & strDescription & "</description>" & vbNewLine

            ' Minimum version of Microsoft controls used with this application
            strXML = strXML & "  <dependency>" & vbNewLine
            strXML = strXML & "    <dependentAssembly>" & vbNewLine
            strXML = strXML & "      <assemblyIdentity" & vbNewLine
            strXML = strXML & "        type=" & Chr$(34) & "win32" & Chr$(34) & vbNewLine
            strXML = strXML & "        name=" & Chr$(34) & "Microsoft.Windows.Common-Controls" & Chr$(34) & vbNewLine
            strXML = strXML & "        version=" & Chr$(34) & "6.0.0.0" & Chr$(34) & vbNewLine
            strXML = strXML & "        processorArchitecture=" & Chr$(34) & "*" & Chr$(34) & vbNewLine
            strXML = strXML & "        publicKeyToken=" & Chr$(34) & "6595b64144ccf1df" & Chr$(34) & vbNewLine
            strXML = strXML & "      />" & vbNewLine
            strXML = strXML & "    </dependentAssembly>" & vbNewLine
            strXML = strXML & "  </dependency>" & vbNewLine

            ' Current operating system
            strXML = strXML & "  <compatibility xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:compatibility.v1" & Chr$(34) & ">" & vbNewLine
            strXML = strXML & "    <application>" & vbNewLine
            strXML = strXML & "      <!-- Windows 10 / Server 2016 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      <!-- Windows 8.1 / Server 2012 R2 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{1f676c76-80e1-4239-95bb-83d0f6d0da78}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      <!-- Windows 8 / Server 2012 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      <!-- Windows 7 / Server 2008 R2 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{35138b9a-5d96-4fbd-8e2d-a2440225f93a}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "      <!-- Windows Vista / Server 2008 -->" & vbNewLine
            strXML = strXML & "      <supportedOS Id=" & Chr$(34) & "{e2011457-1546-43c5-a5fe-008deee3d3f0}" & Chr$(34) & "/>" & vbNewLine
            strXML = strXML & "    </application>" & vbNewLine
            strXML = strXML & "  </compatibility>" & vbNewLine

            ' Only recognized by Windows 7 or newer
            If (mlngMajorVer = 6 And mlngMinorVer > 0) Or _
               (mlngMajorVer = 10) Then

                strXML = strXML & "  <trustInfo xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v3" & Chr$(34) & ">" & vbNewLine
                strXML = strXML & "    <security>" & vbNewLine
                strXML = strXML & "      <requestedPrivileges>" & vbNewLine
                strXML = strXML & "        <requestedExecutionLevel" & vbNewLine
                strXML = strXML & "          level=" & Chr$(34) & mstrSecurityLevel & Chr$(34) & vbNewLine
                strXML = strXML & "          uiAccess=" & Chr$(34) & "false" & Chr$(34) & vbNewLine
                strXML = strXML & "        />" & vbNewLine
                strXML = strXML & "      </requestedPrivileges>" & vbNewLine
                strXML = strXML & "    </security>" & vbNewLine
                strXML = strXML & "  </trustInfo>" & vbNewLine
            End If

            strXML = strXML & "</assembly>"
        Else
            ' Windows XP manifest file
            strXML = strXML & "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & _
                              "UTF-8" & Chr$(34) & " standalone=" & Chr$(34) & "yes" & Chr$(34) & "?>" & vbNewLine

            strXML = strXML & "<assembly xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v1" & Chr$(34) & _
                              " manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & ">" & vbNewLine

            ' Minimum version of this application
            strXML = strXML & "  <assemblyIdentity" & vbNewLine
            strXML = strXML & "    type=" & Chr$(34) & "win32" & Chr$(34) & vbNewLine
            strXML = strXML & "    name=" & Chr$(34) & strExeName & Chr$(34) & vbNewLine
            strXML = strXML & "    version=" & Chr$(34) & strAppVersion & Chr$(34) & vbNewLine
            strXML = strXML & "    processorArchitecture=" & Chr$(34) & "*" & Chr$(34) & vbNewLine
            strXML = strXML & "    publicKeyToken=" & Chr$(34) & "0000000000000000" & Chr$(34) & vbNewLine
            strXML = strXML & "  />" & vbNewLine

            ' Company and application name
            strXML = strXML & "  <description>" & strDescription & "</description>" & vbNewLine

            ' Minimum version of Microsoft controls used with this application
            strXML = strXML & "  <dependency>" & vbNewLine
            strXML = strXML & "    <dependentAssembly>" & vbNewLine
            strXML = strXML & "      <assemblyIdentity" & vbNewLine
            strXML = strXML & "        type=" & Chr$(34) & "win32" & Chr$(34) & vbNewLine
            strXML = strXML & "        name=" & Chr$(34) & "Microsoft.Windows.Common-Controls" & Chr$(34) & vbNewLine
            strXML = strXML & "        version=" & Chr$(34) & "6.0.0.0" & Chr$(34) & vbNewLine
            strXML = strXML & "        processorArchitecture=" & Chr$(34) & "*" & Chr$(34) & vbNewLine
            strXML = strXML & "        publicKeyToken=" & Chr$(34) & "6595b64144ccf1df" & Chr$(34) & vbNewLine
            strXML = strXML & "      />" & vbNewLine
            strXML = strXML & "    </dependentAssembly>" & vbNewLine
            strXML = strXML & "  </dependency>" & vbNewLine
            strXML = strXML & "</assembly>"
        End If

        ' Create a new file or
        ' empty an existing file
        hFile = FreeFile
        Open strFileName For Output As #hFile
        Print #hFile, strXML
        Close #hFile

        CreateManifestFile = True   ' Successful finish

    Else

        ' Manifest file is not needed
        If IsPathValid(strFileName) Then
            Kill strFileName   ' Remove old file
            DoEvents
        End If

    End If

CreateManifestFile_CleanUp:
    If Err.Number <> 0 Then
        Err.Clear
    End If

    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       NeedManifest
'
' Description:   Evaluate current operating system
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Aug-2015  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' **************************************************************************
Private Function NeedManifest(ByRef blnVer6orNewer As Boolean, _
                              ByRef blnWin7orEarlier As Boolean) As Boolean

    ' Called by CreateManifestFile()

    ' Preset flags to FALSE
    NeedManifest = False
    blnVer6orNewer = False
    blnWin7orEarlier = False

    ' If no data passed then
    ' set to lowest values
    If mlngMajorVer = 0 Then
        mlngMajorVer = 5
        mlngMinorVer = 0
    End If

    Select Case mlngMajorVer
           Case 5   ' Win XP
                If mlngMinorVer > 0 Then
                    NeedManifest = True
                End If

           Case 6   ' Win Vista, 7, 8, 8.1, 10 beta
                NeedManifest = True
                blnVer6orNewer = True

                If mlngMinorVer < 2 Then
                    blnWin7orEarlier = True
                End If

           Case 10   ' Win 10
                blnVer6orNewer = True
                NeedManifest = True

           Case Else
    End Select

End Function

' ***************************************************************************
' Routine:       FileVersion
'
' Description:   Capture file version information
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Version information
'
' Example:       "1.0.0.0" <- "C:\Kens Software\ABCD.exe"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2012  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function FileVersion(ByVal strFileName As String) As String

    ' Called by CreateManifestFile()

    Dim lngPointer As Long               ' pointer to version info data
    Dim lngLength  As Long               ' length of info pointed at by lngPointer
    Dim abytData() As Byte               ' Buffer for version info resource
    Dim typFFInfo  As VS_FIXEDFILEINFO   ' File version structure

    strFileName = QualifyPath(App.Path) & strFileName & ".exe"
    lngLength = 0

    FileVersion = ""                       ' Pre-empty return value
    ZeroMemory typFFInfo, Len(typFFInfo)   ' Empty type structure

    ' Get file size information
    lngLength = GetFileVersionInfoSize(strFileName, lngPointer)

    ' No EXE available
    If lngLength = 0 Then

        ' Are we in the VB IDE?
        If mblnVB_IDE Then
            ' Capture major version number from vbp file
            FileVersion = CStr(App.Major) & ".0.0.0"
        Else
            FileVersion = "0.0.0.0"   ' No appl EXE major version available
        End If

        Exit Function   ' Exit routine

    End If

    ReDim abytData(0 To lngLength - 1)   ' Size data array

    ' Get version information
    Call GetFileVersionInfo(strFileName, 0, lngLength, abytData(0))

    ' Get pointer to structure that holds the data
    Call VerQueryValue(abytData(0), "\", lngPointer, lngLength)

    ' Copy data into type structure
    CopyMemory typFFInfo, ByVal lngPointer, lngLength

    ' Capture major version number
    With typFFInfo
        FileVersion = TrimStr(CStr(HIWORD(.dwFileVersionMS))) & ".0.0.0"
    End With

    Erase abytData()   ' Empty array

End Function

Private Function HIWORD(ByVal dwValue As Long) As Long

    ' Called by FileVersion()

    Dim strHex As String

    strHex = Right$("00000000" & Hex$(dwValue), 8)
    HIWORD = CLng("&H" & Left$(strHex, 4))

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
    ' Enum enumSecurityLevel
    Const eMan_Invoker       As Long = 0&
    Const eMan_HighestAvail  As Long = 1&
    Const eMan_Administrator As Long = 2&
End Sub
#End If

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    ' Default property values
    MajorVersion = 5
    MinorVersion = 0
    SecurityLevel = eMan_Invoker
    bVB_IDE = True

End Sub

