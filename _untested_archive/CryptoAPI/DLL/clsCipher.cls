VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsCipher"
  
' ***************************************************************************
' Module Variables
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Byte)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a byt Password
' Variable name:     mabytPassword
' ***************************************************************************
  Private mblnCreateNewFile As Boolean
  Private mlngBlockSize     As Long
  Private mlngKeyLength     As Long
  Private mlngHashMethod    As enumAPI_HashAlgorithms
  Private mlngCipherMethod  As enumCryptoAPI_CipherAlgorithms
  Private mabytPassword()   As Byte
  Private mobjHash          As cHash
  
  ' Define class object with events (See bottom of this module)
  Private WithEvents mobjCryptoAPI  As cCryptoAPI
Attribute mobjCryptoAPI.VB_VarHelpID = -1
  

' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event CipherProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

' Comment out Version Property if not used in this application
Public Property Get Version() As String
    ' Version of the DLL this class is associated with.
    Version = DLL_NAME & ".dll v" & GetFileVerData(App.Path & "\" & DLL_NAME & ".dll") & _
              vbNewLine & " Copyright (c) 2004-2017 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Let StopProcessing(ByVal blnData As Boolean)
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing."
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing."
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property

Public Property Let CreateNewFile(ByVal blnNewFile As Boolean)
    mblnCreateNewFile = blnNewFile
End Property

Public Property Let HashMethod(ByVal lngHashMethod As Long)
    mlngHashMethod = lngHashMethod
End Property

Public Property Let CipherMethod(ByVal lngCipherMethod As Long)
    mlngCipherMethod = lngCipherMethod
End Property

Public Property Let BlockSize(ByVal lngBlockSize As Long)
    mlngBlockSize = lngBlockSize
End Property

Public Property Let KeyLength(ByVal lngKeyLength As Long)
    mlngKeyLength = lngKeyLength
End Property

Public Property Get PasswordLength_Min() As Long
    PasswordLength_Min = MIN_PWD_LENGTH
End Property

Public Property Get PasswordLength_Max() As Long
    PasswordLength_Max = MAX_PWD_LENGTH
End Property

Public Property Get AES_Ready() As Boolean
    AES_Ready = mobjCryptoAPI.AES_Ready
End Property

Public Property Let Password(ByRef abytData() As Byte)

    ' Used when assigning a value to the password property,
    ' on the left side of an assignment.
    '
    ' Syntax:     X.Password = "User Supplied Password"
    '
    ' My recommendation:
    '    The length of the Password/Passphrase should be
    '    a minimum of eight characters (ASCII 32-255) in
    '    order to obtain a good mix.
    '
    Const ROUTINE_NAME As String = "Password Property"

    On Error GoTo Password_Error

    ' Appended null character for last position in byte array.
    ' It will not be acknowledged during processing.
    ReDim mabytPassword(UBound(abytData) + 1)
    
    If UBound(abytData) > 0 Then
        
        CopyMemory mabytPassword(0), abytData(0), UBound(abytData) + 1
        ReDim Preserve mabytPassword(UBound(abytData))
        
        If Not ValidPassword Then
            gblnStopProcessing = True
            Erase mabytPassword()
            DoEvents
        End If
    
    Else
        InfoMsg "Missing or invalid password." & vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        
        gblnStopProcessing = True
        Erase mabytPassword()
        DoEvents
    End If

Password_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Property

Password_Error:
    ErrorMsg MODULE_NAME, "Password Property", Err.Description
    gblnStopProcessing = True
    Resume Password_CleanUp
    
End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

Public Function StringProcessing(ByRef abytData() As Byte, _
                                 ByVal blnEncrypt As Boolean) As Boolean
    
    Dim abytPwd() As Byte

    Const ROUTINE_NAME As String = "StringProcessing"

    On Error GoTo StringProcessing_Error

    StringProcessing = False   ' Preset flag to FALSE
    Erase abytPwd()            ' Always start with an empty array

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' is the data string empty
    If Not CBool(IsArrayInitialized(abytData)) Then
        InfoMsg "The data string is empty.  There is no data to encrypt." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo StringProcessing_CleanUp
    End If
    
    With mobjCryptoAPI
        Select Case mlngCipherMethod
        
               Case 0 To 6
                    ' Validate password
                    If ValidPassword() Then
                    
                        ' It is best to use a byte array to pass data
                        ' and passwords around an application.
                        With mobjHash
                            .StopProcessing = gblnStopProcessing       ' Set flag to contiune processing
                            .HashMethod = mlngHashMethod               ' Set hash method
                            abytPwd() = .HashString(mabytPassword())   ' Hash password
                            gblnStopProcessing = .StopProcessing       ' See if processing has been aborted
                        End With
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo StringProcessing_CleanUp
                        End If
                
                    Else
                        ' An invalid password was used
                        gblnStopProcessing = True
                        GoTo StringProcessing_CleanUp
                    End If
                        
                    .StopProcessing = gblnStopProcessing   ' Set flag to contiune processing
                    .HashMethod = mlngHashMethod           ' Save session hash method
                    .CipherMethod = mlngCipherMethod       ' Save cipher method
                    .BlockSize = mlngBlockSize             ' AES block size (128, 192, 256)
                    .KeyLength = mlngKeyLength             ' AES key length (128, 192, 256)
                    .Password = abytPwd()                  ' Save password
            
                    ' Validate hash and cipher selections
                    If Not .AES_Ready Then
                        Select Case mlngCipherMethod
                               Case 3 To 6
                                    InfoMsg "CryptoAPI RSA AES provider required for" & vbNewLine & _
                                            "this cipher selection is not available." & vbNewLine & _
                                            "Update your selections and try again."
                                    GoTo StringProcessing_CleanUp
                        End Select
                        
                        Select Case mlngHashMethod
                               Case 4 To 6
                                    InfoMsg "CryptoAPI RSA AES provider required for" & vbNewLine & _
                                            "this hash selection is not available." & vbNewLine & _
                                            "Update your selections and try again."
                                    GoTo StringProcessing_CleanUp
                        End Select
                        
                    End If
               
               Case 7   ' Base64
                    .CipherMethod = mlngCipherMethod
        End Select
        
        ' Perform string Encryption/Decryption
        StringProcessing = .CryptoAPI_String(abytData(), blnEncrypt)
        
        gblnStopProcessing = .StopProcessing   ' See if processing has been aborted
    
    End With

StringProcessing_CleanUp:
    Erase abytPwd()   ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

StringProcessing_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume StringProcessing_CleanUp

End Function

Public Function FileProcessing(ByVal strSource As String, _
                               ByVal blnEncrypt As Boolean) As Boolean

    Dim curFileSize As Currency
    Dim abytData()  As Byte
    Dim abytPwd()   As Byte
    Dim objBigFiles As cBigFiles

    Const ROUTINE_NAME As String = "FileProcessing"

    On Error GoTo FileProcessing_Error

    FileProcessing = False   ' Preset flag to FALSE
    Erase abytData()         ' Always start with empty arrays
    Erase abytPwd()
    curFileSize = 0@
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' is the data string empty
    If Len(TrimStr(strSource)) = 0 Then
        InfoMsg "Cannot identify Path\File name." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo FileProcessing_CleanUp
    End If
    
    ' Make sure the source file do exist
    If Not IsPathValid(strSource) Then
        ' the source file does NOT exist
        InfoMsg Chr$(34) & strSource & Chr$(34) & vbNewLine & _
                "cannot be found." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo FileProcessing_CleanUp
    End If
    
    Set objBigFiles = New cBigFiles
    objBigFiles.CalcFileSize strSource, curFileSize
    Set objBigFiles = Nothing
    
    ' Is the file empty?
    If curFileSize < 1 Then
        InfoMsg "File:  " & strSource & vbNewLine & _
                "is empty.  There is no data to process." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo FileProcessing_CleanUp
    End If

    ' Check for enough free space to process this file
    If Not IsEnoughFreeSpace(strSource, curFileSize) Then
        GoTo FileProcessing_CleanUp
    End If
        
    ' Save path/file name to byte array
    abytData() = StringToByteArray(strSource)

    With mobjCryptoAPI
        Select Case mlngCipherMethod
        
               Case 0 To 6
                    ' Validate password
                    If ValidPassword() Then
                    
                        ' It is best to use a byte array to pass data
                        ' and passwords around an application.
                        With mobjHash
                            .StopProcessing = gblnStopProcessing       ' Set flag to contiune processing
                            .HashMethod = mlngHashMethod               ' Set hash method
                            abytPwd() = .HashString(mabytPassword())   ' Hash password
                            gblnStopProcessing = .StopProcessing       ' See if processing has been aborted
                        End With
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo FileProcessing_CleanUp
                        End If
                
                    Else
                        ' An invalid password was used
                        gblnStopProcessing = True
                        GoTo FileProcessing_CleanUp
                    End If
                    
                    .StopProcessing = gblnStopProcessing   ' Set flag to contiune processing
                    .HashMethod = mlngHashMethod           ' Save session hash method
                    .CipherMethod = mlngCipherMethod       ' Save cipher method
                    .BlockSize = mlngBlockSize             ' AES block size (128, 192, 256)
                    .KeyLength = mlngKeyLength             ' AES key length (128, 192, 256)
                    .Password = abytPwd()                  ' Save password
            
                    ' Validate hash and cipher selections
                    If Not .AES_Ready Then
                        Select Case mlngCipherMethod
                               Case 3 To 6
                                    InfoMsg "CryptoAPI RSA AES provider required for" & vbNewLine & _
                                            "this cipher selection is not available." & vbNewLine & _
                                            "Update your selections and try again."
                                    GoTo FileProcessing_CleanUp
                        End Select
                        
                        Select Case mlngHashMethod
                               Case 4 To 6
                                    InfoMsg "CryptoAPI RSA AES provider required for" & vbNewLine & _
                                            "this hash selection is not available." & vbNewLine & _
                                            "Update your selections and try again."
                                    GoTo FileProcessing_CleanUp
                        End Select
                        
                    End If
               
                    ' Perform file Encryption/Decryption
                    FileProcessing = .CryptoAPI_File(abytData(), blnEncrypt)
               
               Case 7   ' Base64
                    .CipherMethod = mlngCipherMethod
               
                    ' Perform file Encryption/Decryption
                    FileProcessing = .Base64_File(abytData(), blnEncrypt)
        End Select
        
        gblnStopProcessing = .StopProcessing   ' See if processing has been aborted
    
    End With
    
FileProcessing_CleanUp:
    Close              ' Close all open files
    Erase abytData()   ' Always empty arrays when not needed
    Erase abytPwd()
    
    Set objBigFiles = Nothing   ' Free class objects from memory
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

FileProcessing_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume FileProcessing_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function ValidPassword(Optional ByVal strPassword As String = "") As Boolean

    ' Called by Property Password()
    '           StringProcessing()
    '           FileProcessing()
    
    Const ROUTINE_NAME As String = "ValidPassword"

    On Error GoTo ValidPassword_Error
         
    ValidPassword = False   ' Preset flag to FALSE
    
    If Len(strPassword) > 0 Then
        
        strPassword = TrimStr(strPassword)   ' Remove unwanted characters
    
        If (Len(strPassword) + 1) < MIN_PWD_LENGTH Then
            InfoMsg "Minimum password length is " & _
                    MIN_PWD_LENGTH & " characters." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
            GoTo ValidPassword_CleanUp
        
        ElseIf (Len(strPassword) + 1) > MAX_PWD_LENGTH Then
            InfoMsg "Maximum password length is " & _
                    MAX_PWD_LENGTH & " characters." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
            GoTo ValidPassword_CleanUp
        End If
    
        ValidPassword = True
        GoTo ValidPassword_CleanUp
        
    End If
         
    If CBool(IsArrayInitialized(mabytPassword())) Then
         
        If (UBound(mabytPassword) + 1) < MIN_PWD_LENGTH Then
            InfoMsg "Minimum password length is " & _
                    MIN_PWD_LENGTH & " characters." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
            GoTo ValidPassword_CleanUp
        
        ElseIf (UBound(mabytPassword) + 1) > MAX_PWD_LENGTH Then
            InfoMsg "Maximum password length is " & _
                    MAX_PWD_LENGTH & " characters." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
            GoTo ValidPassword_CleanUp
        End If
        
    Else
        InfoMsg "Missing or invalid password." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo ValidPassword_CleanUp
    End If
    
    ValidPassword = True

ValidPassword_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

ValidPassword_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ValidPassword_CleanUp
    
End Function

Private Function IsEnoughFreeSpace(ByVal strSource As String, _
                                   ByVal curFileSize As Currency) As Boolean

    Dim strTemp      As String
    Dim strDrive     As String
    Dim curFreeSpace As Currency
    Dim objFSO       As Scripting.FileSystemObject
    Dim objDrive     As Drive

    Const ROUTINE_NAME As String = "IsEnoughFreeSpace"
    Const MB_10        As Currency = 10485760@
    
    On Error GoTo IsEnoughFreeSpace_Error

    IsEnoughFreeSpace = False   ' Preset flag
    
    ' Test for free space on the
    ' system temp folder drive
    strTemp = GetTempFolder()
    strDrive = UCase$(Left$(strTemp, 3))
    
    ' Determine amount of free space for the temp folder drive
    Set objFSO = New Scripting.FileSystemObject
    Set objDrive = objFSO.GetDrive(strDrive)
    curFreeSpace = objDrive.FreeSpace
    
    ' Adjust file size to include a 10mb buffer
    curFileSize = curFileSize + MB_10
    
    ' See if there is enough free space to
    ' hold the new file plus a buffer of
    ' 10mb in the designated temp folder
    If curFileSize >= curFreeSpace Then
            
        InfoMsg "File:  " & strSource & vbNewLine & _
                "File size + buffer:   " & Format$(curFileSize, "#,##0") & vbNewLine & _
                "Disk free space:  " & Format$(curFreeSpace, "#,##0") & vbNewLine & _
                "There is not enough free space on the temp folder drive (" & _
                strDrive & ") to process this file." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    ' Test for free space
    ' on the source drive
    strDrive = UCase$(Left$(strSource, 3))
    Set objDrive = objFSO.GetDrive(strDrive)
    curFreeSpace = objDrive.FreeSpace
    
    ' See if there is enough free space to
    ' hold the new file plus a buffer of
    ' 10mb on the source folder drive
    If curFileSize >= curFreeSpace Then
            
        InfoMsg "File:  " & strSource & vbNewLine & _
                "File size + buffer:   " & Format$(curFileSize, "#,##0") & vbNewLine & _
                "Disk free space:  " & Format$(curFreeSpace, "#,##0") & vbNewLine & _
                "There is not enough free space on the source folder drive (" & _
                strDrive & ") to hold the new file file." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    IsEnoughFreeSpace = True

IsEnoughFreeSpace_CleanUp:
    Set objDrive = Nothing  ' Always empty objects when not needed
    Set objFSO = Nothing
    On Error GoTo 0         ' Nullify this error trap
    Exit Function

IsEnoughFreeSpace_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume IsEnoughFreeSpace_CleanUp
    
End Function

Private Sub Class_Initialize()
         
    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
                
    Erase mabytPassword()         ' Empty password array
    Set mobjHash = Nothing        ' Free class objects from memory
    Set mobjCryptoAPI = Nothing
    
    ' Instantiate class objects
    Set mobjHash = New cHash
    Set mobjCryptoAPI = New cCryptoAPI
    
    ' Preset property values
    StopProcessing = False   ' Set to continue processing
    CipherMethod = 1         ' Set default cipher algorithm
    BlockSize = 128          ' AES block size (128, 192, 256)
    KeyLength = 128          ' AES key length (128, 192, 256)
    HashMethod = 4           ' Set default hash algorithm SHA-256
    CreateNewFile = True     ' Do not overlay file being encrypted
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    Erase mabytPassword()         ' Empty password array
    Set mobjHash = Nothing        ' Free class objects from memory
    Set mobjCryptoAPI = Nothing
    
End Sub


' ***************************************************************************
' ****                      Events from subclasses                       ****
' ***************************************************************************
    
' Progress event is raised from the appropriate
' class and then raised from here to the form.

Private Sub mobjCryptoAPI_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub

