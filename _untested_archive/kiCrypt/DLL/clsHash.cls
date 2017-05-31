VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cHash  (clsHash.cls)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' 31-Dec-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added property HashRounds()
' 02-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated references to RipeMD128_256 and RipeMD160_320.
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Add events to track hash progress on the screen.
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd Mar-2012 (Supercedes FIPS 180-3)
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Combined all Tiger hashes into a single class.
'              - Removed reference to Tiger2 class.
' 10-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Used clsAPI_Hash class to replace cMD4, cMD5, cSHA1 and
'                cSHA2 classes.
'              - Removed RipeMD classes because they are considered weak.
' 04-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added new property ReturnLowercse().  Designates if returned
'              hashed data string should be in upper or lowercase format.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME  As String = "clsHash"
  
' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumHASH_ALGORITHM
      eHASH_MD2             ' 0
      eHASH_MD4             ' 1
      eHASH_MD5             ' 2
      eHASH_SHA1            ' 3
      eHASH_SHA256          ' 4
      eHASH_SHA384          ' 5
      eHASH_SHA512          ' 6
      eHASH_Tiger3_128      ' 7
      eHASH_Tiger3_160      ' 8
      eHASH_Tiger3_192      ' 9
      eHASH_Tiger3_224      ' 10
      eHASH_Tiger3_256      ' 11
      eHASH_Tiger3_384      ' 12
      eHASH_Tiger3_512      ' 13
      eHASH_Whirlpool_224   ' 14
      eHASH_Whirlpool_256   ' 15
      eHASH_Whirlpool_384   ' 16
      eHASH_Whirlpool_512   ' 17
  End Enum
  
' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Long)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m lng HashMethod
' Variable name:     mlngHashMethod
'
' ***************************************************************************
  Private mlngHashRounds   As Long
  Private mblnRetLowercase As Boolean
  Private mlngHashMethod   As enumHASH_ALGORITHM
    
  ' Define class objects with events (See bottom of this module)
  Private WithEvents mobjAPI_Hash  As cAPI_Hash
Attribute mobjAPI_Hash.VB_VarHelpID = -1
  Private WithEvents mobjTiger3    As cTiger3
Attribute mobjTiger3.VB_VarHelpID = -1
  Private WithEvents mobjWhirlpool As cWhirlpool
Attribute mobjWhirlpool.VB_VarHelpID = -1
  

' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


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

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = True
    mblnRetLowercase = blnData
End Property

Public Property Let HashMethod(ByVal lngHashMethod As enumHASH_ALGORITHM)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eHASH_SHA256

    Select Case mlngHashMethod
           Case 0 To 17: mlngHashMethod = lngHashMethod
           Case Else:    mlngHashMethod = 3   ' eAPI_SHA1
    End Select
    
End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1

    Select Case mlngHashMethod
    
           ' MD2, MD4, MD5, SHA family, Whirlpool family
           Case 0 To 6, 14 To 17
                Select Case lngData
                       Case 1 To 10   ' Good choice
                            mlngHashRounds = lngData
                       Case Else      ' Use default value
                            mlngHashRounds = 1
                End Select

           Case 7 To 13    ' Tiger3 family
                Select Case lngData
                       Case 3 To 15
                            mlngHashRounds = lngData   ' Good choice
                       Case Else
                            mlngHashRounds = 3         ' Use default value
                End Select
    End Select
    
End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       HashString
'
' Description:   Creates a hash output string based on string data input.
'
' Parameters:    abytData() - Data string in byte array to be hashed
'
'    Example:    abytData() = StrConv("abc", vbFromUnicode)
'
' Returns:       Formatted hashed data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 20-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array.  Added functionality
'              to return data in various formats.
' 04-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in output format.
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Combined all Tiger hashes into a single class.
' ***************************************************************************
Public Function HashString(ByRef abytData() As Byte) As Variant

    Dim abytHash() As Byte
    
    On Error GoTo HashString_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If
    
    Erase abytHash()  ' Always start with empty arrays
    ResetHashObjects  ' Verify hash objects are empty

    Select Case mlngHashMethod

           Case eHASH_MD2, eHASH_MD4, eHASH_MD5, eHASH_SHA1, _
                eHASH_SHA256, eHASH_SHA384, eHASH_SHA512
                
                Set mobjAPI_Hash = New cAPI_Hash    ' Instantiate class object
                With mobjAPI_Hash
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_MD2:    .HashMethod = 0
                           Case eHASH_MD4:    .HashMethod = 1
                           Case eHASH_MD5:    .HashMethod = 2
                           Case eHASH_SHA1:   .HashMethod = 3
                           Case eHASH_SHA256: .HashMethod = 4
                           Case eHASH_SHA384: .HashMethod = 5
                           Case eHASH_SHA512: .HashMethod = 6
                    End Select
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_Tiger3_128, eHASH_Tiger3_160, eHASH_Tiger3_192, _
                eHASH_Tiger3_224, eHASH_Tiger3_256, eHASH_Tiger3_384, _
                eHASH_Tiger3_512
                
                Set mobjTiger3 = New cTiger3   ' Instantiate class object
                With mobjTiger3
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_Tiger3_128: .HashMethod = 0
                           Case eHASH_Tiger3_160: .HashMethod = 1
                           Case eHASH_Tiger3_192: .HashMethod = 2
                           Case eHASH_Tiger3_224: .HashMethod = 3
                           Case eHASH_Tiger3_256: .HashMethod = 4
                           Case eHASH_Tiger3_384: .HashMethod = 5
                           Case eHASH_Tiger3_512: .HashMethod = 6
                    End Select
                    abytHash() = .HashString(abytData())
                End With

           Case eHASH_Whirlpool_224, eHASH_Whirlpool_256, _
                eHASH_Whirlpool_384, eHASH_Whirlpool_512
                
                Set mobjWhirlpool = New cWhirlpool   ' Instantiate class object
                With mobjWhirlpool
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_Whirlpool_224: .HashMethod = 0
                           Case eHASH_Whirlpool_256: .HashMethod = 1
                           Case eHASH_Whirlpool_384: .HashMethod = 2
                           Case eHASH_Whirlpool_512: .HashMethod = 3
                    End Select
                    abytHash() = .HashString(abytData())
                End With
    End Select

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ReDim abytHash(0)  ' Resize to smallest number of elements
    End If

    HashString = abytHash()
           
HashString_CleanUp:
    Erase abytHash()  ' Always empty arrays when not needed
    ResetHashObjects  ' Verify hash objects are empty
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

HashString_Error:
    ErrorMsg MODULE_NAME, "HashString", Err.Description
    Resume HashString_CleanUp

End Function

' ***************************************************************************
' Routine:       HashFile
'
' Description:   Creates a hashed data string based on file contents.
'
' Parameters:    abytData() - Data file location in byte array format
'
'                Ex:  abytData() = StrConv("C:\Test Folder\Testfile.doc", vbFromUnicode)
'
' Returns:       Formatted hashed data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 20-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array.  Added functionality
'              to return data in various formats.
' 04-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in output format.
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Combined all Tiger hashes into a single class.
' ***************************************************************************
Public Function HashFile(ByRef abytData() As Byte) As Variant

    Dim abytHash() As Byte
    
    On Error GoTo HashFile_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If
    
    Erase abytHash()  ' Always start with empty arrays
    ResetHashObjects  ' Verify hash objects are empty

    Select Case mlngHashMethod

           Case eHASH_MD2, eHASH_MD4, eHASH_MD5, eHASH_SHA1, _
                eHASH_SHA256, eHASH_SHA384, eHASH_SHA512
                
                Set mobjAPI_Hash = New cAPI_Hash    ' Instantiate class object
                With mobjAPI_Hash
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_MD2:    .HashMethod = 0
                           Case eHASH_MD4:    .HashMethod = 1
                           Case eHASH_MD5:    .HashMethod = 2
                           Case eHASH_SHA1:   .HashMethod = 3
                           Case eHASH_SHA256: .HashMethod = 4
                           Case eHASH_SHA384: .HashMethod = 5
                           Case eHASH_SHA512: .HashMethod = 6
                    End Select
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_Tiger3_128, eHASH_Tiger3_160, eHASH_Tiger3_192, _
                eHASH_Tiger3_224, eHASH_Tiger3_256, eHASH_Tiger3_384, _
                eHASH_Tiger3_512
                
                Set mobjTiger3 = New cTiger3   ' Instantiate class object
                With mobjTiger3
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_Tiger3_128: .HashMethod = 0
                           Case eHASH_Tiger3_160: .HashMethod = 1
                           Case eHASH_Tiger3_192: .HashMethod = 2
                           Case eHASH_Tiger3_224: .HashMethod = 3
                           Case eHASH_Tiger3_256: .HashMethod = 4
                           Case eHASH_Tiger3_384: .HashMethod = 5
                           Case eHASH_Tiger3_512: .HashMethod = 6
                    End Select
                    abytHash() = .HashFile(abytData())
                End With

           Case eHASH_Whirlpool_224, eHASH_Whirlpool_256, _
                eHASH_Whirlpool_384, eHASH_Whirlpool_512
                
                Set mobjWhirlpool = New cWhirlpool   ' Instantiate class object
                With mobjWhirlpool
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    Select Case mlngHashMethod
                           Case eHASH_Whirlpool_224: .HashMethod = 0
                           Case eHASH_Whirlpool_256: .HashMethod = 1
                           Case eHASH_Whirlpool_384: .HashMethod = 2
                           Case eHASH_Whirlpool_512: .HashMethod = 3
                    End Select
                    abytHash() = .HashFile(abytData())
                End With
    End Select

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ReDim abytHash(0)  ' Resize to smallest number of elements
    End If

    HashFile = abytHash()
           
HashFile_CleanUp:
    Erase abytHash()  ' Always empty arrays when not needed
    ResetHashObjects  ' Verify hash objects are empty
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

HashFile_Error:
    ErrorMsg MODULE_NAME, "HashFile", Err.Description
    Resume HashFile_CleanUp

End Function

Private Sub ResetHashObjects()

    ' Free class objects from memory
    Set mobjAPI_Hash = Nothing
    Set mobjTiger3 = Nothing
    Set mobjWhirlpool = Nothing

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
    ' Enum enumHASH_ALGORITHM
    Const eHASH_MD2           As Long = 0&
    Const eHASH_MD4           As Long = 1&
    Const eHASH_MD5           As Long = 2&
    Const eHASH_SHA1          As Long = 3&
    Const eHASH_SHA256        As Long = 4&
    Const eHASH_SHA384        As Long = 5&
    Const eHASH_SHA512        As Long = 6&
    Const eHASH_Tiger3_128    As Long = 7&
    Const eHASH_Tiger3_160    As Long = 8&
    Const eHASH_Tiger3_192    As Long = 9&
    Const eHASH_Tiger3_224    As Long = 10&
    Const eHASH_Tiger3_256    As Long = 11&
    Const eHASH_Tiger3_384    As Long = 12&
    Const eHASH_Tiger3_512    As Long = 13&
    Const eHASH_Whirlpool_224 As Long = 14&
    Const eHASH_Whirlpool_256 As Long = 15&
    Const eHASH_Whirlpool_384 As Long = 16&
    Const eHASH_Whirlpool_512 As Long = 17&
End Sub
#End If

Private Sub Class_Initialize()
    
    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
                
    ResetHashObjects  ' Verify hash objects are empty
    
    ' Preset property values
    ReturnLowercase = True   ' Default return of hashed data
    StopProcessing = False   ' Set to continue processing
    HashMethod = 3           ' Set default hash algorithm (eAPI_SHA1)
    HashRounds = 1           ' Set default number of passes

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    ResetHashObjects

End Sub


' ***************************************************************************
' ****                      Events from subclasses                       ****
' ***************************************************************************
    
' Progress event is raised from the appropriate
' class and then raised from here to the form.

Private Sub mobjAPI_Hash_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjTiger3_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjWhirlpool_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub


