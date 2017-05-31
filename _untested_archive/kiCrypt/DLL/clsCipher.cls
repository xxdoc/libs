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
  Private Const MB_5        As Long = &H500000   ' 5242880 bytes
  
' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumCIPHER_ALGORITHM
      eCIPHER_ARCFOUR     ' 0
      eCIPHER_BASE64      ' 1
      eCIPHER_BLOWFISH    ' 2
      eCIPHER_GOST        ' 3
      eCIPHER_RIJNDAEL    ' 4
      eCIPHER_SERPENT     ' 5
      eCIPHER_SKIPJACK    ' 6
      eCIPHER_TWOFISH     ' 7
  End Enum

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
  Private mobjHash          As cHash
  Private mobjMath32        As cMath32
  Private mlngCipherMethod  As enumCIPHER_ALGORITHM
  Private mlngHashMethod    As enumHASH_ALGORITHM
  Private mblnCreateNewFile As Boolean
  Private mlngRounds        As Long
  Private mlngKeyRounds     As Long
  Private mlngKeyLength     As Long
  Private mlngBlockSize     As Long
  Private mabytPassword()   As Byte
  
  ' Define class objects with events (See bottom of this module)
  Private WithEvents mobjArcFour  As cArcFour
Attribute mobjArcFour.VB_VarHelpID = -1
  Private WithEvents mobjBase64   As cBase64
Attribute mobjBase64.VB_VarHelpID = -1
  Private WithEvents mobjBlowfish As cBlowFish
Attribute mobjBlowfish.VB_VarHelpID = -1
  Private WithEvents mobjGost     As cGost
Attribute mobjGost.VB_VarHelpID = -1
  Private WithEvents mobjRijndael As cRijndael
Attribute mobjRijndael.VB_VarHelpID = -1
  Private WithEvents mobjSerpent  As cSerpent
Attribute mobjSerpent.VB_VarHelpID = -1
  Private WithEvents mobjSkipjack As cSkipjack
Attribute mobjSkipjack.VB_VarHelpID = -1
  Private WithEvents mobjTwofish  As cTwofish
Attribute mobjTwofish.VB_VarHelpID = -1
  

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
    Version = DLL_NAME & ".dll (tm) v" & App.Major & "." & App.Minor & "." & App.Revision & _
              " Copyright (c) 2004-2017 - Kenneth Ives  kenaso@tx.rr.com"
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

Public Property Let KeyLength(ByVal lngKeyLength As Long)
    mlngKeyLength = lngKeyLength
End Property

Public Property Let PrimaryKeyRounds(ByVal lngKeyRounds As Long)
    ' Only used by Blowfish and GOST
    mlngKeyRounds = lngKeyRounds
End Property

Public Property Let BlockSize(ByVal lngBlockSize As Long)
    mlngBlockSize = lngBlockSize
End Property

Public Property Let CipherRounds(ByVal lngRounds As Long)
    mlngRounds = lngRounds
End Property

Public Property Let HashMethod(ByVal lngHashMethod As enumHASH_ALGORITHM)
    mlngHashMethod = lngHashMethod
End Property

Public Property Let CipherMethod(ByVal lngCipherMethod As enumCIPHER_ALGORITHM)
    mlngCipherMethod = lngCipherMethod
End Property

Public Property Get PasswordLength_Min() As Long
    PasswordLength_Min = MIN_PWD_LENGTH
End Property

Public Property Get PasswordLength_Max() As Long
    PasswordLength_Max = MAX_PWD_LENGTH
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
            Erase mabytPassword()
            gblnStopProcessing = True
            DoEvents
        End If
        
    Else
        InfoMsg "Missing or invalid password." & vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        
        Erase mabytPassword()
        gblnStopProcessing = True
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

Public Function EncryptString(ByRef abytData() As Byte) As Boolean
    
    Dim abytPwd() As Byte

    Const ROUTINE_NAME As String = "EncryptString"

    On Error GoTo EncryptString_Error

    EncryptString = False   ' Preset flag to FALSE
    Erase abytPwd()         ' Always start with an empty array

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' is the data string empty
    If Not CBool(IsArrayInitialized(abytData)) Then
        InfoMsg "No data to encrypt." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo EncryptString_CleanUp
    End If
    
    If mlngCipherMethod = eCIPHER_BASE64 Then
    
        ReDim abytPwd(0)  ' dummy password just to use as a parameter
            
    Else
        If ValidPassword() Then
    
            Erase abytPwd()   ' empty temp password array
            
            ' It is best to use a byte array to pass data
            ' and passwords around an application.
            With mobjHash
                .HashMethod = mlngHashMethod              ' set the hash method
                abytPwd() = .HashString(mabytPassword())  ' hash the password
            End With
        
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo EncryptString_CleanUp
            End If

        Else
            ' An invalid password was used
            gblnStopProcessing = True
            GoTo EncryptString_CleanUp
        End If
        
    End If
    
    ' perform the encryption process
    If ProcessString(abytData(), abytPwd(), True) Then
        EncryptString = True
    End If
    
EncryptString_CleanUp:
    Erase abytPwd()   ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

EncryptString_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptString_CleanUp

End Function

Public Function DecryptString(ByRef abytData() As Byte) As Boolean

    Dim abytPwd() As Byte

    Const ROUTINE_NAME As String = "DecryptString"

    On Error GoTo DecryptString_Error

    DecryptString = False   ' Preset flag to FALSE
    Erase abytPwd()         ' Always start with an empty array

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' is the data string empty
    If Not CBool(IsArrayInitialized(abytData)) Then
        InfoMsg "No data to decrypt." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo DecryptString_CleanUp
    End If
    
    If mlngCipherMethod = eCIPHER_BASE64 Then
    
        ReDim abytPwd(0)   ' dummy password just to use as a parameter

    Else
        If ValidPassword() Then
    
            Erase abytPwd()   ' empty temp password array
            
            ' It is best to use a byte array to pass data
            ' and passwords around an application.
            With mobjHash
                .HashMethod = mlngHashMethod              ' set the hash method
                abytPwd() = .HashString(mabytPassword())  ' hash the password
            End With
        
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo DecryptString_CleanUp
            End If

        Else
            ' An invalid password was used
            gblnStopProcessing = True
            GoTo DecryptString_CleanUp
        End If
    
    End If
    
    ' perform the decryption process
    If ProcessString(abytData(), abytPwd(), False) Then
        DecryptString = True
    End If

DecryptString_CleanUp:
    Erase abytPwd()   ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

DecryptString_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptString_CleanUp

End Function

Public Function EncryptFile(ByVal strSource As String) As Boolean

    Dim curFileSize As Currency
    Dim abytData()  As Byte
    Dim abytPwd()   As Byte
    Dim objBigFiles As cBigFiles

    Const ROUTINE_NAME As String = "EncryptFile"

    On Error GoTo EncryptFile_Error

    EncryptFile = False  ' Preset flag to FALSE
    Erase abytData()     ' Always start with empty arrays
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
        GoTo EncryptFile_CleanUp
    End If
    
    ' Make sure the source file do exist
    If IsPathValid(strSource) Then
        
        Set objBigFiles = New cBigFiles
        objBigFiles.CalcFileSize strSource, curFileSize
        Set objBigFiles = Nothing
        
        ' Is the file empty?
        If curFileSize < 1 Then
            InfoMsg "File:  " & strSource & vbNewLine & _
                    "is empty.  There is no data to process." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
            GoTo EncryptFile_CleanUp
        End If

        ' Check for enough free space to process this file
        If Not IsEnoughFreeSpace(strSource, curFileSize) Then
            GoTo EncryptFile_CleanUp
        End If
            
        ' Save path/file name to byte array
        abytData() = StringToByteArray(strSource)

        ' Process the password
        If mlngCipherMethod = eCIPHER_BASE64 Then
        
            ' Is the file too big?  A Base64 encrypted file
            ' is approximately one third larger than the
            ' original and may exceed limits.
            '
            ' ex:  6973030 = MB_5 + Fix(MB_5 * 0.33)
            '      6973030 = 5242880 + 1730150
            If curFileSize > CCur(MB_5 + Fix(MB_5 * 0.33)) Then
                
                InfoMsg "File:  " & strSource & vbNewLine & vbNewLine & _
                        "File size exceeds " & Format$(MB_5, "#,##0") & _
                        " bytes after encryption." & vbNewLine & _
                        "Base64 is not designed to process large files." & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                
                GoTo EncryptFile_CleanUp
            End If

            ReDim abytPwd(0)   ' dummy password just to use as a parameter
        
        Else
            ' Validate the password
            If ValidPassword() Then
            
                ' It is best to use a byte array to pass data
                ' and passwords around an application.
                With mobjHash
                    .HashMethod = mlngHashMethod               ' set the hash method
                    abytPwd() = .HashString(mabytPassword())   ' hash the password
                End With
            
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    GoTo EncryptFile_CleanUp
                End If

            Else
                ' An invalid password was used
                gblnStopProcessing = True
                GoTo EncryptFile_CleanUp
            End If
        End If
        
        ' Perform the encryption process
        If ProcessFile(abytData(), abytPwd(), True) Then
            EncryptFile = True
        End If
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo EncryptFile_CleanUp
        End If

    Else
        ' the source file does NOT exist
        InfoMsg Chr$(34) & strSource & Chr$(34) & vbNewLine & _
                "cannot be found." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
    End If
    
EncryptFile_CleanUp:
    Close              ' close all open files
    Erase abytData()   ' Always empty arrays when not needed
    Erase abytPwd()
    Set objBigFiles = Nothing   ' Free class object from memory

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

EncryptFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptFile_CleanUp

End Function

Public Function DecryptFile(ByVal strSource As String) As Boolean

    Dim curFileSize As Currency
    Dim abytData()  As Byte
    Dim abytPwd()   As Byte
    Dim objBigFiles As cBigFiles

    Const ROUTINE_NAME As String = "DecryptFile"

    On Error GoTo DecryptFile_Error

    DecryptFile = False   ' Preset flag to FALSE
    Erase abytData()      ' Always start with empty arrays
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
        GoTo DecryptFile_CleanUp
    End If
    
    ' Make sure the source file do exist
    If IsPathValid(strSource) Then
        
        Set objBigFiles = New cBigFiles
        objBigFiles.CalcFileSize strSource, curFileSize
        Set objBigFiles = Nothing
        
        ' Is the file empty?
        If curFileSize < 1 Then
            InfoMsg "File:  " & strSource & vbNewLine & _
                    "is empty.  There is no data to decrypt." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
            Exit Function
        End If

        ' Check for enough free space to process this file
        If Not IsEnoughFreeSpace(strSource, curFileSize) Then
            GoTo DecryptFile_CleanUp
        End If
            
        ' Save path/file name to byte array
        abytData() = StringToByteArray(strSource)

        ' Process the password
        If mlngCipherMethod = eCIPHER_BASE64 Then
        
            ' Is the file too big?  A Base64 encrypted file
            ' is approximately one third larger than the
            ' original and may exceed limits.
            '
            ' ex:  6973030 = MB_5 + Fix(MB_5 * 0.33)
            '      6973030 = 5242880 + 1730150
            If curFileSize > CCur(MB_5 + Fix(MB_5 * 0.33)) Then
                InfoMsg "File:  " & strSource & vbNewLine & _
                        "File size exceeds the maximum limit set for Base64." & vbNewLine & _
                        "Base64 is not designed to process large files." & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
                GoTo DecryptFile_CleanUp
            End If

            ReDim abytPwd(0)   ' dummy password just to use as a parameter
        
        Else
            ' Validate the password
            If ValidPassword() Then
            
                ' It is best to use a byte array to pass data
                ' and passwords around an application.
                With mobjHash
                    .HashMethod = mlngHashMethod               ' set the hash method
                    abytPwd() = .HashString(mabytPassword())   ' hash the password
                End With
            
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    GoTo DecryptFile_CleanUp
                End If

            Else
                ' An invalid password was used
                gblnStopProcessing = True
                GoTo DecryptFile_CleanUp
            End If
        End If
        
        ' Perform the decryption process
        If ProcessFile(abytData(), abytPwd(), False) Then
            DecryptFile = True
        End If
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo DecryptFile_CleanUp
        End If

    Else
        ' the source file does NOT exist
        InfoMsg Chr$(34) & strSource & Chr$(34) & vbNewLine & _
                "cannot be found." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
    End If
    
DecryptFile_CleanUp:
    Close              ' Close all open files
    Erase abytData()   ' Always empty arrays when not needed
    Erase abytPwd()
    Set objBigFiles = Nothing   ' Free class object from memory
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

DecryptFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptFile_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function ProcessFile(ByRef abytData() As Byte, _
                             ByRef abytPwd() As Byte, _
                             ByVal blnEncrypt As Boolean) As Boolean

    ' Called by EncryptFile()
    '           DecryptFile()
    
    Const ROUTINE_NAME As String = "ProcessFile"

    On Error GoTo ProcessFile_Error
    
    ProcessFile = True  ' Preset flag
    ResetCipherObjects  ' Verify cipher objects are empty
    
    ' determine which cipher algorithm to use
    Select Case mlngCipherMethod

           Case eCIPHER_ARCFOUR
                Set mobjArcFour = New cArcFour   ' Instantiate class object
                With mobjArcFour
                    .CreateNewFile = mblnCreateNewFile
                    .KeyLength = mlngKeyLength
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessFile = .ArcFour_File(abytData(), blnEncrypt)
                End With

           Case eCIPHER_BASE64
                Set mobjBase64 = New cBase64   ' Instantiate class object
                mobjBase64.CreateNewFile = mblnCreateNewFile  ' 15-Jun-2013 Line was missing
                ProcessFile = mobjBase64.Base64_File(abytData(), blnEncrypt)

           Case eCIPHER_BLOWFISH
                Set mobjBlowfish = New cBlowFish   ' Instantiate class object
                With mobjBlowfish
                    .CreateNewFile = mblnCreateNewFile
                    .KeyLength = mlngKeyLength
                    .PrimaryKeyRounds = mlngKeyRounds
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessFile = .Blowfish_File(abytData(), blnEncrypt)
                End With

           Case eCIPHER_GOST
                Set mobjGost = New cGost   ' Instantiate class object
                With mobjGost
                    .CreateNewFile = mblnCreateNewFile
                    .KeyLength = mlngKeyLength
                    .PrimaryKeyRounds = mlngKeyRounds
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessFile = .GOST_File(abytData(), blnEncrypt)
                End With

           Case eCIPHER_RIJNDAEL
                Set mobjRijndael = New cRijndael   ' Instantiate class object
                With mobjRijndael
                    .CreateNewFile = mblnCreateNewFile
                    .BlockSize = mlngBlockSize
                    .KeyLength = mlngKeyLength
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessFile = .Rijndael_File(abytData(), blnEncrypt)
                End With

           Case eCIPHER_SERPENT
                Set mobjSerpent = New cSerpent   ' Instantiate class object
                With mobjSerpent
                    .CreateNewFile = mblnCreateNewFile
                    .KeyLength = mlngKeyLength
                    .PrimaryKeyRounds = mlngKeyRounds
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessFile = .Serpent_File(abytData(), blnEncrypt)
                End With

           Case eCIPHER_SKIPJACK
                Set mobjSkipjack = New cSkipjack   ' Instantiate class object
                With mobjSkipjack
                    .CreateNewFile = mblnCreateNewFile
                    .KeyLength = mlngKeyLength
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessFile = .Skipjack_File(abytData(), blnEncrypt)
                End With

           Case eCIPHER_TWOFISH
                Set mobjTwofish = New cTwofish   ' Instantiate class object
                With mobjTwofish
                    .CreateNewFile = mblnCreateNewFile
                    .KeyLength = mlngKeyLength
                    .PrimaryKeyRounds = mlngKeyRounds
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessFile = .Twofish_File(abytData(), blnEncrypt)
                End With
    End Select

ProcessFile_CleanUp:
    ResetCipherObjects                    ' Verify cipher objects are empty
    StopProcessing = gblnStopProcessing   ' See if processing has been stopped
    On Error GoTo 0                       ' Nullify this error trap
    Exit Function

ProcessFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ProcessFile = False
    Resume ProcessFile_CleanUp

End Function

Private Function ProcessString(ByRef abytData() As Byte, _
                               ByRef abytPwd() As Byte, _
                               ByVal blnEncrypt As Boolean) As Boolean

    ' Called by EncryptString()
    '           DecryptString()
    
    Const ROUTINE_NAME As String = "ProcessString"

    On Error GoTo ProcessString_Error
    
    ProcessString = True  ' Preset flag
    ResetCipherObjects    ' Verify cipher objects are empty
    
    ' determine which cipher algorithm to use
    Select Case mlngCipherMethod

           Case eCIPHER_ARCFOUR
                Set mobjArcFour = New cArcFour   ' Instantiate class object
                With mobjArcFour
                    .KeyLength = mlngKeyLength
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessString = .ArcFour_String(abytData(), blnEncrypt)
                End With

           Case eCIPHER_BASE64
                Set mobjBase64 = New cBase64   ' Instantiate class object
                ProcessString = mobjBase64.Base64_String(abytData(), blnEncrypt)

           Case eCIPHER_BLOWFISH
                Set mobjBlowfish = New cBlowFish   ' Instantiate class object
                With mobjBlowfish
                    .KeyLength = mlngKeyLength
                    .PrimaryKeyRounds = mlngKeyRounds
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessString = .Blowfish_String(abytData(), blnEncrypt)
                End With

           Case eCIPHER_GOST
                Set mobjGost = New cGost   ' Instantiate class object
                With mobjGost
                    .KeyLength = mlngKeyLength
                    .PrimaryKeyRounds = mlngKeyRounds
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessString = .GOST_String(abytData(), blnEncrypt)
                End With

           Case eCIPHER_RIJNDAEL
                Set mobjRijndael = New cRijndael   ' Instantiate class object
                With mobjRijndael
                    .BlockSize = mlngBlockSize
                    .KeyLength = mlngKeyLength
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessString = .Rijndael_String(abytData(), blnEncrypt)
                End With

           Case eCIPHER_SERPENT
                Set mobjSerpent = New cSerpent   ' Instantiate class object
                With mobjSerpent
                    .KeyLength = mlngKeyLength
                    .PrimaryKeyRounds = mlngKeyRounds
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessString = .Serpent_String(abytData(), blnEncrypt)
                End With

           Case eCIPHER_SKIPJACK
                Set mobjSkipjack = New cSkipjack   ' Instantiate class object
                With mobjSkipjack
                    .KeyLength = mlngKeyLength
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessString = .Skipjack_String(abytData(), blnEncrypt)
                End With

           Case eCIPHER_TWOFISH
                Set mobjTwofish = New cTwofish   ' Instantiate class object
                With mobjTwofish
                    .KeyLength = mlngKeyLength
                    .PrimaryKeyRounds = mlngKeyRounds
                    .EncryptRounds = mlngRounds
                    .Password = abytPwd()
                    ProcessString = .Twofish_String(abytData(), blnEncrypt)
                End With
    End Select

ProcessString_CleanUp:
    ResetCipherObjects                    ' Verify cipher objects are empty
    StopProcessing = gblnStopProcessing   ' See if processing has been stopped
    On Error GoTo 0                       ' Nullify this error trap
    Exit Function

ProcessString_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ProcessString = False
    Resume ProcessString_CleanUp

End Function

Private Function ValidPassword(Optional ByVal strPassword As String = "") As Boolean

    ' Called by EncryptFile()
    '           DecryptFile()
    '           EncryptString()
    '           DecryptString()
    '           Property Password()
    
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

    ' Called by EncryptFile()
    '           DecryptFile()
    
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

Private Sub ResetCipherObjects()

    ' Free class objects from memory
    Set mobjGost = Nothing
    Set mobjBase64 = Nothing
    Set mobjArcFour = Nothing
    Set mobjSerpent = Nothing
    Set mobjTwofish = Nothing
    Set mobjBlowfish = Nothing
    Set mobjRijndael = Nothing
    Set mobjSkipjack = Nothing
    
End Sub

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
    ' Enum enumCIPHER_ALGORITHM
    Const eCIPHER_ARCFOUR  As Long =  0&
    Const eCIPHER_BASE64   As Long =  1&
    Const eCIPHER_BLOWFISH As Long =  2&
    Const eCIPHER_GOST     As Long =  3&
    Const eCIPHER_RIJNDAEL As Long =  4&
    Const eCIPHER_SERPENT  As Long =  5&
    Const eCIPHER_SKIPJACK As Long =  6&
    Const eCIPHER_TWOFISH  As Long =  7&
End Sub        
#End If

Private Sub Class_Initialize()
         
    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
                
    Erase mabytPassword()
    ResetCipherObjects            ' Verify cipher objects are empty
    
    Set mobjHash = Nothing        ' Free class objects from memory
    Set mobjMath32 = Nothing
    
    Set mobjHash = New cHash      ' Instantiate class objects
    Set mobjMath32 = New cMath32
    
    ' Default values
    StopProcessing = False
    KeyLength = 128
    BlockSize = 128
    CipherRounds = 1
    CipherMethod = eCIPHER_BLOWFISH
    HashMethod = eHASH_SHA256
    CreateNewFile = True
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    Erase mabytPassword()
    ResetCipherObjects        ' Verify cipher objects are empty
    
    Set mobjHash = Nothing    ' Free class objects from memory
    Set mobjMath32 = Nothing

End Sub


' ***************************************************************************
' ****                      Events from subclasses                       ****
' ***************************************************************************
    
' Progress event is raised from the appropriate
' class and then raised from here to the form.

Private Sub mobjArcFour_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub

Private Sub mobjBase64_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub

Private Sub mobjBlowfish_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub

Private Sub mobjGost_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub

Private Sub mobjRijndael_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub

Private Sub mobjSerpent_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub

Private Sub mobjSkipjack_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub

Private Sub mobjTwofish_CipherProgress(ByVal lngProgress As Long)
    RaiseEvent CipherProgress(lngProgress)
End Sub
