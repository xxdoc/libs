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
' 31-May-2015  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME  As String = "clsHash"
  
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
  Private mlngHashMethod   As enumAPI_HashAlgorithms
    
  ' Define class objects with events (See bottom of this module)
  Private WithEvents mobjAPI_Hash  As cAPI_Hash
Attribute mobjAPI_Hash.VB_VarHelpID = -1
  

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

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = True
    mblnRetLowercase = blnData
End Property

Public Property Let HashMethod(ByVal lngHashMethod As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eAPI_SHA256

    Select Case mlngHashMethod
    
           ' MD2, MD4, MD5, SHA2 family
           Case 0 To 6: mlngHashMethod = lngHashMethod
           Case Else:   mlngHashMethod = 3   ' eAPI_SHA1
    End Select
    
End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1

    Select Case mlngHashMethod
    
           ' MD2, MD4, MD5, SHA2 family
           Case 0 To 6: mlngHashRounds = lngData
           Case Else:   mlngHashRounds = 1
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
    
    Erase abytHash()   ' Always empty arrays when not needed
    
    With mobjAPI_Hash
        Select Case mlngHashMethod
    
               Case eAPI_MD2, eAPI_MD4, eAPI_MD5, eAPI_SHA1, _
                    eAPI_SHA256, eAPI_SHA384, eAPI_SHA512
                    
                    Select Case mlngHashMethod
                           Case eAPI_MD2:    .HashMethod = 0
                           Case eAPI_MD4:    .HashMethod = 1
                           Case eAPI_MD5:    .HashMethod = 2
                           Case eAPI_SHA1:   .HashMethod = 3
                           Case eAPI_SHA256: .HashMethod = 4
                           Case eAPI_SHA384: .HashMethod = 5
                           Case eAPI_SHA512: .HashMethod = 6
                    End Select
                    
               Case Else: .HashMethod = eAPI_SHA1   ' 3
        End Select
    
        .HashRounds = mlngHashRounds
        .ReturnLowercase = mblnRetLowercase
        abytHash() = .HashString(abytData())
    
    End With
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ReDim abytHash(0)  ' Resize to smallest number of elements
    End If

    HashString = abytHash()
           
HashString_CleanUp:
    Erase abytHash()   ' Always empty arrays when not needed
    On Error GoTo 0    ' Nullify this error trap
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
    
    Erase abytHash()   ' Always empty arrays when not needed
    
    With mobjAPI_Hash
        Select Case mlngHashMethod
    
               Case eAPI_MD2, eAPI_MD4, eAPI_MD5, eAPI_SHA1, _
                    eAPI_SHA256, eAPI_SHA384, eAPI_SHA512
                    
                    Select Case mlngHashMethod
                           Case eAPI_MD2:    .HashMethod = 0
                           Case eAPI_MD4:    .HashMethod = 1
                           Case eAPI_MD5:    .HashMethod = 2
                           Case eAPI_SHA1:   .HashMethod = 3
                           Case eAPI_SHA256: .HashMethod = 4
                           Case eAPI_SHA384: .HashMethod = 5
                           Case eAPI_SHA512: .HashMethod = 6
                    End Select
                    
               Case Else: .HashMethod = 3   ' eAPI_SHA1
        End Select
    
        .HashRounds = mlngHashRounds
        .ReturnLowercase = mblnRetLowercase
        abytHash() = .HashString(abytData())
    
    End With
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ReDim abytHash(0)  ' Resize to smallest number of elements
    End If

    HashFile = abytHash()
           
HashFile_CleanUp:
    Erase abytHash()   ' Always empty arrays when not needed
    On Error GoTo 0    ' Nullify this error trap
    Exit Function

HashFile_Error:
    ErrorMsg MODULE_NAME, "HashFile", Err.Description
    Resume HashFile_CleanUp

End Function

Private Sub Class_Initialize()
    
    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
                
    Set mobjAPI_Hash = Nothing         ' Verify class object is not active
    Set mobjAPI_Hash = New cAPI_Hash   ' Instantiate class object
    
    ' Preset property values
    ReturnLowercase = True   ' Default return of hashed data
    StopProcessing = False   ' Set to continue processing
    HashMethod = 3           ' Set default hash algorithm (eAPI_SHA1)
    HashRounds = 1           ' Set default number of passes

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    Set mobjAPI_Hash = Nothing   ' Free class object from memory

End Sub


' ***************************************************************************
' ****                      Events from subclasses                       ****
' ***************************************************************************
    
' Progress event is raised from the appropriate
' class and then raised from here to the form.

Private Sub mobjAPI_Hash_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub


