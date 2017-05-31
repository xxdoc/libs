VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTiger3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:       cTiger3  (clsTiger3.cls)
'
' This module is my experimental version for Tiger-128 thru Tiger-512 bit
' output.  I call it Tiger3.  I have not had any problems with it thus far.
'
' I decided to make changes to the initial work data array [malngHash()].
' I believe that this is a cleaner and more secure method as to the
' calculation of the Tiger hashes even tho the results do not match the
' test vectors on the Tiger2 web page.  See LoadWorkArrays() routine for
' more details.  This module was written using a combination of the 32-bit
' and 64-bit code, therefore, none of the test vectors will match.  Most
' of the calculation functions are 32-bit.
'
' If you should encounter any problems, please email me at:
'
'           Kenneth Ives  kenaso@tx.rr.com
'
' DO NOT CONTACT   Ross Anderson, http://www.cl.cam.ac.uk/users/rja14/
'                  Eli Biham, http://www.cs.technion.ac.il/~biham
'
' because they did not write this module nor are they responsible in any
' manner as to its content or output.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
' *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING ***
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
'    You acknowledge that this software is subject to the export control
'    laws and regulations of the United States ("U.S.") and agree to abide
'    by those laws and regulations. Under U.S. law, this software may not
'    be downloaded or otherwise exported, reexported, or transferred to
'    restricted countries, restricted end-users, or for restricted
'    end-uses. The U.S. currently has embargo restrictions against Cuba,
'    Iran, Iraq, Libya, North Korea, Sudan, and Syria. The lists of
'    restricted end-users are maintained on the U.S. Commerce Department's
'    Denied Persons List, the Commerce Department's Entity List, the
'    Commerce Department's List of Unverified Persons, and the U.S.
'    Treasury Department's List of Specially Designated Nationals and
'    Blocked Persons. In addition, this software may not be downloaded or
'    otherwise exported, reexported, or transferred to an end-user engaged
'    in activities related to weapons of mass destruction.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
' Description:   Tiger2 is a fast New cHash function, designed to be very fast
'                on modern computers, and in particular on the state-of-the-art
'                64-bit computers (like DEC-Alpha), while it is still not
'                slower than other suggested hash functions on 32-bit machines.
'
'                Tiger hash has no usage restrictions nor patents. It can be
'                used freely, with the reference implementation, with other
'                implementations or with a modification to the reference
'                implementation (as long as it still implements Tiger2).  The
'                authors only ask that you to let them know about your
'                implementation and to cite the origin of Tiger and of the
'                reference implementation.
'                http://www.cs.technion.ac.il/~biham/Reports/Tiger/
'
' My opinion:    Tiger2 is a very strong hash.  Unfortunately, the 32-bit
'                version seems to be just an off shoot and is not the primary
'                concern of the authors.  That is understandable because
'                everyone appears to be headed for 64-bit hash scenarios.
'                The test vectors, I believe, were created prior to the new
'                MD5 padding of Tiger2 and may be obsolete.  The time stamp
'                on the 32-bit version of the Tiger code is dated 30-Mar-1996
'                while the code was updated, I believe, in 2007.  When I
'                contacted one of the authors, the response was less than
'                enthusiastic concerning the 32-bit version of Tiger2.
'
' Reference:     Tiger2 - A Fast New Hash function
'                http://www.cs.technion.ac.il/~biham/Reports/Tiger/
'
'                Original authors of the Tiger hash:
'                   Ross Anderson, http://www.cl.cam.ac.uk/users/rja14/
'                   Eli Biham, http://www.cs.technion.ac.il/~biham
'
'                Tiger2 64-bit Source code dtd 8-Feb-1996
'                http://www.cs.technion.ac.il/~biham/Reports/Tiger/tiger-src.zip
'
'                Tiger2 32-bit C source code dtd 30-Mar-1996
'                http://www.cs.technion.ac.il/~biham/Reports/Tiger/tiger-src32.zip
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-SEP-2003  Kenneth Ives  kenaso@tx.rr.com
' 26-MAR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Removed my experimental versions until a later date.
' 02-MAY-2006  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote this module to use only long integers.  Much faster.
'              Found C source code at:
'              http://www.cs.technion.ac.il/~biham/Reports/Tiger/tiger-src32.zip
' 22-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              ** MD5 padding - The data is stored in Little-Endian format
'                 with the Least Significant Bit (LSB) first.  Message data is
'                 converted to BIG-ENDIAN format prior to manipulation and then
'                 back to LITTLE-ENDIAN when formatting final output.
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 07-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified output to 224, 256, 384, 512 bits.
' 27-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified Transform() routine for a better mix of data.
' 31-Dec-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Added property HashRounds() to set number of rounds for hashing.
'              - Added property HashMethod() to determine hash algorithm.
'              - Add hashed results to original hash values as last
'                calculation in Transform() routine for a better mix of data.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while hashing data.
'                See Transform() routine.
'              - Removed obsolete code.  Updated documentation.
' 23-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated Transform() routine for speed
' 31-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Modified the way Tiger-224, Tiger-256, Tiger-384, Tiger-512
'                calculates hashes by creating unique initialization values
'                for each.  See LoadWorkArrays() routine for more details.
'              - Updated documentation.
' 19-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated reference documentation
'              - Fixed logic bug concerning flow of events within the number
'                of passes being performed.  See Tranform() routine
' 07-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in KeySchedule() routine. One of the shift values
'              was 19 now is 9
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Combined all Tiger hashes into a single module
'              - Created unique initialization values for each of the Tiger
'                hashes except Tiger-192
'              - Kept original values for constant arrays
' 23-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed a bug in KeySchedule() routine.  Corrected two of the
'                index pointers
'              - Added some more comments
' 04-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added new property ReturnLowercse().  Designates if returned
'              hashed data string should be in upper or lowercase format.
' 23-May-2013  Kenneth Ives  kenaso@tx.rr.com
'              Major rewrite of this module to more closely resemble
'              original code.  All 64-Bit string code was way too slow while
'              all 32-bit long integer code did not feel secure.
' 28-Dec-2014  Kenneth Ives  kenaso@tx.rr.com
'              - Updated CalcRound() routine which changed output data
'              - Updated Documentation
' 02-Feb-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Updated Shift64() with a flag to also perform bit rotation
'              - Updated CalcRound() routine to use bit rotation
'              - Updated Transform() routine to add new hashes to malngHash()
'                array starting with last element of HH() hash array
'              - Fixed bug in FormatOutput() routine, added "Step 2" to loop
' 08-Aug-2016  Kenneth Ives  kenaso@tx.rr.com
'              - Removed Shift64() routine.  Now referencing clsMath64.cls.
'              - Test results have been updated to reflect new calculations.
' 05-Sep-2016  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote PrepareMsgData(), KeySchedule() routines.
'              - Increased block size capture from 32768 to 5,000,000 bytes.
'              - Created unique initialization values for each of
'                the Tiger3 hashes.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsTiger3"
  Private Const MAX_ARRAY   As Long = 256        ' Max array size
  Private Const MB_5        As Long = &H500000   ' 5242880 bytes
  
' ***************************************************************************
' Enumerations
' ***************************************************************************
  Private Enum enumTIGER3_ALGORITHM
      eTiger3_128   ' 0
      eTiger3_160   ' 1
      eTiger3_192   ' 2
      eTiger3_224   ' 3
      eTiger3_256   ' 4
      eTiger3_384   ' 5
      eTiger3_512   ' 6
  End Enum

' ***************************************************************************
' Module Variables
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Long)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a lng Hash
' Variable name:     malngHash
'
' ***************************************************************************
  Private mblnRetLowercase As Boolean
  Private mcurMaxSize      As Currency
  Private mcurAccrued      As Currency
  Private mlngMaxBytes     As Long
  Private mlngSpaceAlloc   As Long
  Private mlngHashRounds   As Long
  Private T1()             As Long   ' Constants tables
  Private T2()             As Long
  Private T3()             As Long
  Private T4()             As Long
  Private malngHash()      As Long
  Private mobjMath32       As cMath32
  Private mobjMath64       As cMath64
  Private mlngHashMethod   As enumTIGER3_ALGORITHM


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let HashRounds(ByVal lngData As Long)

    ' Three (3) passes are minimum for hashing the data.
    ' Use four or more passes when you need extra security.
    Select Case lngData
           Case 3 To 25: mlngHashRounds = lngData   ' Good choice
           Case Else:    mlngHashRounds = 3         ' Use default value
    End Select

End Property

Public Property Let HashMethod(ByVal lngData As Long)

    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eTiger3_192
    Select Case lngData
           
           ' eTiger3_128, eTiger3_160, eTiger3_192, eTiger3_224
           ' eTiger3_256, eTiger3_384, eTiger3_512
           Case 0 To 6
                mlngHashMethod = lngData
                
                Select Case mlngHashMethod
                       Case eTiger3_128, eTiger3_160, eTiger3_192
                            mlngSpaceAlloc = 128
                            mlngMaxBytes = 64
            
                       Case eTiger3_224, eTiger3_256, eTiger3_384, eTiger3_512
                            mlngSpaceAlloc = 256
                            mlngMaxBytes = 128
                End Select

           Case Else
                mlngHashMethod = 2   ' eTiger3_192 (Default)
    End Select

End Property

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = False
    mblnRetLowercase = blnData
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       HashFile
'
' Description:   Function to quickly convert a file into a hex string
'
' Syntax:        With mobjTiger3
'                    ' convert file location to byte array 
'                    abytData() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)
'                    ' hash data and return as Byte array
'                    mabytHashed() = .HashFile(abytData())
'                    ' convert byte array to string data
'                    txtData.Text = StrConv(mabytHashed(), vbUnicode)
'                End With
' 
' Parameters:    abytData() - name of file to be hashed in byte array
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Adjusted mcurMaxSize so progress bar display would be close
'              to same hash finish time
' ***************************************************************************
Public Function HashFile(ByRef abytFile() As Byte) As Byte()

    Dim blnEndOfData As Boolean
    Dim hFile        As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim strSource    As String
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    On Error GoTo HashFile_CleanUp

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    Set objBigFiles = New cBigFiles  ' Instantiate class object

    LoadWorkArrays         ' load the work and constant arrays
    mcurAccrued = 0@       ' Init accrued progress
    curFilePos = 0@        ' Set to first position in the file
    Erase abytData()
    blnEndOfData = False   ' Set flag to FALSE
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    With objBigFiles
        strSource = ByteArrayToString(abytFile())   ' Convert byte array to string
        .CalcFileSize strSource, curAmtLeft         ' Get size of file

        ' Calculate amount (double for hex chars)
        ' plus number of hashing rounds requested
        mcurMaxSize = (curAmtLeft * 2) * mlngHashRounds

        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile) Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        Do
            Erase abytData()                              ' Empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)       ' Process the source file in chunks
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process
            
            If curAmtLeft < 1 Then
                blnEndOfData = True   ' No more data to process after this chunk
            End If
            
            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)

            ' Process the message data into the appropriate
            ' string length and perform the bit manipulations.
            If Not PrepareMsgData(abytData(), blnEndOfData) Then
                gblnStopProcessing = True
            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do  ' exit Do..Loop
            End If

        Loop

        .API_CloseFile hFile    ' Close all files opened by this routine
    End With

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    ' Concatenate the state array into one data
    ' string and return the digital signature.
    HashFile = FormatOutput()

HashFile_CleanUp:
    objBigFiles.API_CloseFile hFile  ' Close all files opened by this routine
    Erase malngHash()                ' Empty arrays when not needed
    Erase abytData()
    Set objBigFiles = Nothing        ' Always free objects from memory

End Function

' ***************************************************************************
' Routine:       HashString
'
' Description:   Function to digest a text string and output the result as a
'                string of hexadecimal characters.
'
' Syntax:        With mobjTiger3
'                    ' convert file location to byte array 
'                    abytData() = StrConv("abc", vbFromUnicode)
'                    ' hash data and return as Byte array
'                    mabytHashed() = .HashString(abytData())
'                    ' convert byte array to string data
'                    txtData.Text = StrConv(mabytHashed(), vbUnicode)
'                End With
' 
' Parameters:    abytData() - data to be hashed in byte array
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Adjusted mcurMaxSize so progress bar display would be close
'              to same hash finish time
' ***************************************************************************
Public Function HashString(ByRef abytData() As Byte) As Byte()

    On Error GoTo HashString_CleanUp

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

    LoadWorkArrays                             ' load work and constant arrays
    mcurAccrued = 0@                           ' Init accrued progress
    mcurMaxSize = CCur(UBound(abytData) + 1)   ' Capture data size

    ' Calculate amount with number
    ' of hashing rounds requested
    mcurMaxSize = mcurMaxSize * mlngHashRounds

    ' Process the message data into the appropriate
    ' string length and perform the bit manipulations.
    If PrepareMsgData(abytData()) Then

        ' Concatenate hashed data into a string
        HashString = FormatOutput()
    Else
        gblnStopProcessing = True
    End If

HashString_CleanUp:
    Erase malngHash()   ' Empty arrays when not needed

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

' ***************************************************************************
' Routine:       FormatOutput
'
' Description:   Concatenate hashed values into one string.
'
'                Output length in hex format:
'                    TIGER-128 =  32 characters
'                    TIGER-160 =  40 characters
'                    TIGER-192 =  48 characters
'                    TIGER-224 =  56 characters
'                    TIGER-256 =  64 characters
'                    TIGER-384 =  96 characters
'                    TIGER-512 = 128 characters
'
' Returns:       Concatenated hashed data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-DEC-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Combined all Tiger hashes
' ***************************************************************************
Private Function FormatOutput() As Byte()

    Dim lngIndex   As Long
    Dim strTemp1   As String
    Dim strTemp2   As String
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim strOutput  As String

    strOutput = Space$(150)   ' Preload output string with spaces
    lngPointer = 1            ' Starting position in output string

    ' Determine output string length
    ' based on two character hex hash
    Select Case mlngHashMethod
           Case eTiger3_128: lngLength = 32
           Case eTiger3_160: lngLength = 40
           Case eTiger3_192: lngLength = 48
           Case eTiger3_224: lngLength = 56
           Case eTiger3_256: lngLength = 64
           Case eTiger3_384: lngLength = 96
           Case eTiger3_512: lngLength = 128
    End Select

    With mobjMath32
        For lngIndex = 0 To (UBound(malngHash) - 1) Step 2

            ' Convert long integers to hex
            strTemp1 = Hex$(malngHash(lngIndex))
            strTemp2 = Hex$(malngHash(lngIndex + 1))

            ' All hashed data is converted from BIG_ENDIAN
            ' format to LITTLE_ENDIAN format as performed
            ' in MD5 hashing to include any prefix zero
            ' padding.
            ' Ex:  SwapEndian(AB4512, 8) -> 001245AB
            '
            ' Concatenate Hash values into one
            ' string using alternate positions
            Mid$(strOutput, lngPointer, 16) = .SwapEndian(strTemp2, 8) & _
                                              .SwapEndian(strTemp1, 8)
            lngPointer = lngPointer + 16   ' Increment output pointer

        Next lngIndex
    End With

    ' Remove unwanted leading/trailing blanks
    strOutput = TrimStr(strOutput)
    
    ' Adjust output string accordingly
    strOutput = Left$(strOutput, lngLength)

    ' Convert output data to upper or lower case
    If mblnRetLowercase Then
        strOutput = LCase$(strOutput)
    Else
        strOutput = UCase$(strOutput)
    End If

    FormatOutput = StringToByteArray(strOutput)   ' Convert string data to byte array
    RaiseEvent HashProgress(100)                  ' Update progress bar

End Function

' *************************************************************************** 
' Routine:       PrepareMsgData
' 
' Description:   All data is stored in Big_Endian format with Most 
'                Significant Bit (MSB) first.  In the original version of
'                Tiger, at the end of the data, 0x01 was replaced with 0x80
'                and filled with zeroes to end of string.  Then it became
'                known as Tiger2.
' 
' Parameters:    abytData() - message to be hashed in a byte array 
' 
' Returns:       TRUE if process completed successfully.
'                FALSE if user reqested to stop processing or an error
'                      occurred.
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 05-Sep-2016  Kenneth Ives  kenaso@tx.rr.com 
'              Rewrote routine 
' *************************************************************************** 
Private Function PrepareMsgData(ByRef abytData() As Byte, _
                       Optional ByVal blnEndOfData As Boolean = True) As Boolean

    Dim strHex     As String
    Dim lngIdx     As Long
    Dim lngIndex   As Long
    Dim lngUBound  As Long
    Dim lngPadding As Long

    Const ROUTINE_NAME As String = "PrepareMsgData"

    On Error GoTo PrepareMsgData_Error

    PrepareMsgData = True          ' Preset flag to TRUE
    lngUBound = UBound(abytData)   ' Capture amount of data passed
    
    Do While lngUBound > 0
    
        strHex = Space$(mlngSpaceAlloc)   ' Preload hex string with blanks spaces
        lngIdx = 1                        ' Set starting position in hex string
        
        ' See if amount of data to process
        ' is less than max amount needed
        If lngUBound < mlngMaxBytes Then
                        
            ' This is last of input data
            For lngIndex = 0 To UBound(abytData)
                Mid$(strHex, lngIdx, 2) = Right$("0" & Hex$(abytData(lngIndex)), 2)
                lngIdx = lngIdx + 2
            Next lngIndex
                    
            strHex = Trim$(strHex) & Hex$(128)          ' Append &H80 to end of data
            lngPadding = mlngSpaceAlloc - Len(strHex)   ' Calc amount of space left in string
                        
            If lngPadding > 0 Then
                strHex = strHex & String$(lngPadding, "0")   ' Fill to end of string with zeroes
            End If
            
            strHex = mobjMath32.SwapEndian(strHex, Len(strHex))   ' Manipulate data string
            Transform strHex                                      ' Hash data string
        
        ' Amount of data equals max length of
        ' data needed and there is no more data
        ElseIf (lngUBound = mlngMaxBytes) And blnEndOfData Then
                            
            ' Convert data to hex string
            For lngIndex = 0 To UBound(abytData)
                Mid$(strHex, lngIdx, 2) = Right$("0" & Hex$(abytData(lngIndex)), 2)
                lngIdx = lngIdx + 2
            Next lngIndex
        
            ' *** String 1 ***
            ' Hex string is exact length needed.
            strHex = mobjMath32.SwapEndian(strHex, Len(strHex))   ' Manipulate data string
            Transform strHex                                      ' Hash data string

            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit Do
            End If

            ' *** String 2 ***
            strHex = Trim$(strHex) & Hex$(128) & String$(mlngSpaceAlloc - 2, "0")   ' Append &H80 + string of zeroes
            strHex = mobjMath32.SwapEndian(strHex, Len(strHex))                     ' Manipulate data string
            Transform strHex                                                        ' Hash data string
        
        ' Plenty of data to process
        Else
            ' convert data to hex string
            For lngIndex = 0 To (mlngMaxBytes - 1)
                Mid$(strHex, lngIdx, 2) = Right$("0" & Hex$(abytData(lngIndex)), 2)
                lngIdx = lngIdx + 2
            Next lngIndex
        
            strHex = mobjMath32.SwapEndian(strHex, Len(strHex))   ' Manipulate data string
            Transform strHex                                      ' Hash data string
        
        End If
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit Do
        End If

        lngUBound = lngUBound - mlngMaxBytes   ' Update amount of data left
        
        If lngUBound > 0 Then
            ReDim Preserve abytData(lngUBound)     ' Resize data array accordingly
        Else
            Erase abytData()   ' Empty array - no more data
        End If
        
    Loop
                
PrepareMsgData_CleanUp:
    strHex = vbNullString   ' Verify string variables are empty

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        PrepareMsgData = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

PrepareMsgData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume PrepareMsgData_CleanUp

End Function

' ***************************************************************************
' Routine:       Transform
'
' Description:   Perform the TIGER main hashing functions here.
'
' Parameters:    strData - Message chunk in single string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote this routine
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Changed the input parameter to byte array
' 27-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified the first looping action to more thoroughly mix
'              the data
' 02-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Add hashed results to original hash values as last calculation
'              for a better mix of data.
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added RaiseEvents to track hash progress
' ***************************************************************************
Private Sub Transform(ByVal strData As String)

    Dim strHex      As String
    Dim lngIdx      As Long
    Dim lngLoop     As Long
    Dim lngTemp     As Long
    Dim lngLoops    As Long
    Dim lngRound    As Long
    Dim lngPointer  As Long
    Dim lngProgress As Long
    Dim HH()        As Long
    Dim alngMsg()   As Long

    ' Determine work array sizes and loop iterations
    Select Case mlngHashMethod
           Case eTiger3_128, eTiger3_160, eTiger3_192
                ReDim HH(6)
                ReDim alngMsg(16)
                lngLoops = 1
                
           Case eTiger3_224, eTiger3_256
                ReDim HH(12)
                ReDim alngMsg(32)
                lngLoops = 2
                
           Case eTiger3_384, eTiger3_512
                ReDim HH(18)
                ReDim alngMsg(32)
                lngLoops = 3
    End Select

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    lngIdx = 0   ' Set index pointer

    ' Load data into an array
    For lngPointer = 1 To Len(strData) Step 8

        strHex = Mid$(strData, lngPointer, 8)   ' Capture eight characters
        alngMsg(lngIdx) = CLng("&H" & strHex)   ' Convert to long integer
        lngIdx = lngIdx + 1                     ' Increment index pointer

    Next lngPointer

    ' #define save_abc \
    '      aa0 = a0; \
    '      aa1 = a1; \
    '      bb0 = b0; \
    '      bb1 = b1; \
    '      cc0 = c0; \
    '      cc1 = c1;
    SaveData HH()

    lngIdx = 0   ' Reset index pointer

    For lngLoop = 1 To lngLoops

        ' pass(a0,a1,b0,b1,c0,c1,5); \
        Pass HH(0 + lngIdx), HH(1 + lngIdx), HH(2 + lngIdx), _
             HH(3 + lngIdx), HH(4 + lngIdx), HH(5 + lngIdx), _
             alngMsg(), 5

        ' key_schedule; \
        KeySchedule alngMsg()

        ' pass(c0,c1,a0,a1,b0,b1,7); \
        Pass HH(4 + lngIdx), HH(5 + lngIdx), HH(0 + lngIdx), _
             HH(1 + lngIdx), HH(2 + lngIdx), HH(3 + lngIdx), _
             alngMsg(), 7

        ' key_schedule; \
        KeySchedule alngMsg()

        ' pass(b0,b1,c0,c1,a0,a1,9); \
        Pass HH(2 + lngIdx), HH(3 + lngIdx), HH(4 + lngIdx), _
             HH(5 + lngIdx), HH(0 + lngIdx), HH(1 + lngIdx), _
             alngMsg(), 9

        lngIdx = lngIdx + 6   ' Increment index pointer

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For   ' exit FOR..NEXT loop
        End If

    Next lngLoop
    
    ' Three (3) passes are minimum for hashing data.
    ' Use four or more passes when you need extra security.
    '
    ' for(pass_no=3; pass_no<PASSES; pass_no++) { \
    '     key_schedule \
    '     pass(a0,a1,b0,b1,c0,c1,9); \
    '     tmpa=a0; a0=c0; c0=b0; b0=tmpa; \
    '     tmpa=a1; a1=c1; c1=b1; b1=tmpa;} \
    For lngRound = 1 To mlngHashRounds

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For   ' exit FOR..NEXT loop
        End If

        lngIdx = 0   ' Reset index pointer
        
        For lngLoop = 1 To lngLoops

            ' key_schedule \
            KeySchedule alngMsg()

            ' pass(a0,a1,b0,b1,c0,c1,9); \
            Pass HH(0 + lngIdx), HH(1 + lngIdx), HH(2 + lngIdx), _
                 HH(3 + lngIdx), HH(4 + lngIdx), HH(5 + lngIdx), _
                 alngMsg(), 9

            ' tmpa=a0; a0=c0; c0=b0; b0=tmpa; \
            lngTemp = HH(0 + lngIdx)
            HH(0 + lngIdx) = HH(4 + lngIdx)
            HH(4 + lngIdx) = HH(2 + lngIdx)
            HH(2 + lngIdx) = lngTemp

            ' tmpa=a1; a1=c1; c1=b1; b1=tmpa;} \
            lngTemp = HH(1 + lngIdx)
            HH(1 + lngIdx) = HH(5 + lngIdx)
            HH(5 + lngIdx) = HH(3 + lngIdx)
            HH(3 + lngIdx) = lngTemp

            lngIdx = lngIdx + 6   ' Increment index pointer

            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit For   ' exit FOR..NEXT loop
            End If

        Next lngLoop
    Next lngRound
    
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Copy hashed data back
    ' to original array
    FeedForward HH()

    ' Update progress bar
    mcurAccrued = mcurAccrued + (Len(strData) * mlngHashRounds)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent HashProgress(lngProgress)

End Sub

Private Sub SaveData(ByRef HH() As Long)

    ' Set inital values for the hash operators
    '
    ' #define save_abc \
    '      aa0 = a0; \
    '      aa1 = a1; \
    '      bb0 = b0; \
    '      bb1 = b1; \
    '      cc0 = c0; \
    '      cc1 = c1;
    Dim lngIndex As Long

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Copy base hex data to temp work array
    For lngIndex = 0 To UBound(malngHash) - 1
        HH(lngIndex) = malngHash(lngIndex)
    Next lngIndex

End Sub

Private Sub FeedForward(ByRef HH() As Long)

    ' Process hashed data and copy
    ' back to original array

    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngLoops As Long

    lngIdx = 0   ' Initialize index pointer

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Loop 1 - will process elements 0-5    (Tiger128, 160, 192)
    ' Loop 2 - will process elements 6-11   (Tiger224, 256)
    ' Loop 3 - will process elements 12-17  (Tiger384, 512)
    Select Case mlngHashMethod
           Case eTiger3_128, eTiger3_160, eTiger3_192: lngLoops = 1
           Case eTiger3_224, eTiger3_256:              lngLoops = 2
           Case eTiger3_384, eTiger3_512:              lngLoops = 3
    End Select

    ' #define feedforward \
    '      xor64(a0, a1, aa0, aa1); \
    '      sub64(b0, b1, bb0, bb1); \
    '      add64(c0, c1, cc0, cc1);
    For lngLoop = 1 To lngLoops

        ' 24-Aug-2015 Fixed logic to copy data to original array
        Xor64 malngHash(0 + lngIdx), malngHash(1 + lngIdx), HH(0 + lngIdx), HH(1 + lngIdx)
        Sub64 malngHash(2 + lngIdx), malngHash(3 + lngIdx), HH(2 + lngIdx), HH(3 + lngIdx)
        Add64 malngHash(4 + lngIdx), malngHash(5 + lngIdx), HH(4 + lngIdx), HH(5 + lngIdx)
        
        lngIdx = lngIdx + 6   ' Increment index pointer

    Next lngLoop

End Sub

' ***************************************************************************
'                     Bit manipulation routines
' ***************************************************************************
' #define pass(a0,a1,b0,b1,c0,c1,mul) \
' a0,a1,b0,b1,c0,c1 values are manipulated and returned
Private Sub Pass(ByRef a0 As Long, _
                 ByRef a1 As Long, _
                 ByRef b0 As Long, _
                 ByRef b1 As Long, _
                 ByRef c0 As Long, _
                 ByRef c1 As Long, _
                 ByRef alngMsg() As Long, _
                 ByVal lngMultiplier As Long)

    ' Called by Transform()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngLoops As Long

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Sub
    End If

    Select Case mlngHashMethod
           Case eTiger3_128, eTiger3_160, eTiger3_192:              lngLoops = 1
           Case eTiger3_224, eTiger3_256, eTiger3_384, eTiger3_512: lngLoops = 2
    End Select

    lngIdx = 0   ' Initialize index pointer

    ' round(a0,a1,b0,b1,c0,c1,x00,x01,mul); \
    ' round(b0,b1,c0,c1,a0,a1,x10,x11,mul); \
    ' round(c0,c1,a0,a1,b0,b1,x20,x21,mul); \
    ' round(a0,a1,b0,b1,c0,c1,x30,x31,mul); \
    ' round(b0,b1,c0,c1,a0,a1,x40,x41,mul); \
    ' round(c0,c1,a0,a1,b0,b1,x50,x51,mul); \
    ' round(a0,a1,b0,b1,c0,c1,x60,x61,mul); \
    ' round(b0,b1,c0,c1,a0,a1,x70,x71,mul);
    For lngLoop = 1 To lngLoops

        CalcRound a0, a1, b0, b1, c0, c1, alngMsg(0 + lngIdx), alngMsg(1 + lngIdx), lngMultiplier
        CalcRound b0, b1, c0, c1, a0, a1, alngMsg(2 + lngIdx), alngMsg(3 + lngIdx), lngMultiplier
        CalcRound c0, c1, a0, a1, b0, b1, alngMsg(4 + lngIdx), alngMsg(5 + lngIdx), lngMultiplier
        CalcRound a0, a1, b0, b1, c0, c1, alngMsg(6 + lngIdx), alngMsg(7 + lngIdx), lngMultiplier
        CalcRound b0, b1, c0, c1, a0, a1, alngMsg(8 + lngIdx), alngMsg(9 + lngIdx), lngMultiplier
        CalcRound c0, c1, a0, a1, b0, b1, alngMsg(10 + lngIdx), alngMsg(11 + lngIdx), lngMultiplier
        CalcRound a0, a1, b0, b1, c0, c1, alngMsg(12 + lngIdx), alngMsg(13 + lngIdx), lngMultiplier
        CalcRound b0, b1, c0, c1, a0, a1, alngMsg(14 + lngIdx), alngMsg(15 + lngIdx), lngMultiplier

        lngIdx = lngIdx + 16   ' Increment index pointer

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For   ' exit FOR..NEXT loop
        End If

    Next lngLoop

End Sub

' #define round(a0,a1,b0,b1,c0,c1,x0,x1,mul) \
' a0,a1,b0,b1,c0,c1 values are manipulated and returned
Private Sub CalcRound(ByRef a0 As Long, _
                      ByRef a1 As Long, _
                      ByRef b0 As Long, _
                      ByRef b1 As Long, _
                      ByRef c0 As Long, _
                      ByRef c1 As Long, _
                      ByVal lngMsg0 As Long, _
                      ByVal lngMsg1 As Long, _
                      ByVal lngMultiplier As Long)

    ' Called by Pass()
    
    Dim lngTemp0 As Long
    Dim lngTemp1 As Long
    
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Sub
    End If

    With mobjMath64
        ' xor64(c0,c1,x0,x1); \
        Xor64 c0, c1, lngMsg0, lngMsg1
        
        ' 02-Feb-2015 Changed shift (>>) to Rotate (>>>)
        ' a -= t1[((c)>>>(0*8))&0xFF] ^ t2[((c)>>>(2*8))&0xFF] ^ \
        '      t3[((c)>>>(4*8))&0xFF] ^ t4[((c)>>>(6*8))&0xFF] ; \
        lngTemp0 = T1(0, (c0 And &HFF&))                 ' No bit movement
        lngTemp1 = T1(1, (c1 And &HFF&))
        
        .Rotate_64 c0, c1, -16                           ' Rotate right 16 positions
        lngTemp0 = lngTemp0 Xor T2(0, (c0 And &HFF&))
        lngTemp1 = lngTemp1 Xor T2(1, (c1 And &HFF&))
        
        .Rotate_64 c0, c1, -32                           ' Rotate right 32 positions
        lngTemp0 = lngTemp0 Xor T3(0, (c0 And &HFF&))
        lngTemp1 = lngTemp1 Xor T3(1, (c1 And &HFF&))
        
        .Rotate_64 c0, c1, -48                           ' Rotate right 48 positions
        lngTemp0 = lngTemp0 Xor T4(0, (c0 And &HFF&))
        lngTemp1 = lngTemp1 Xor T4(1, (c1 And &HFF&))
        
        Sub64 a0, a1, lngTemp0, lngTemp1                 ' Update a0 & a1 using subtraction
    
        ' 02-Feb-2015 Changed shift (>>) to Rotate (>>>)
        ' b += t4[((c)>>>(1*8))&0xFF] ^ t3[((c)>>>(3*8))&0xFF] ^ \
        '      t2[((c)>>>(5*8))&0xFF] ^ t1[((c)>>>(7*8))&0xFF] ; \
        .Rotate_64 c0, c1, -8                            ' Rotate right 8 positions
        lngTemp0 = T4(0, (c0 And &HFF&))
        lngTemp1 = T4(1, (c1 And &HFF&))
        
        .Rotate_64 c0, c1, -24                           ' Rotate right 24 positions
        lngTemp0 = lngTemp0 Xor T3(0, (c0 And &HFF&))
        lngTemp1 = lngTemp1 Xor T3(1, (c1 And &HFF&))
        
        .Rotate_64 c0, c1, -40                           ' Rotate right 40 positions
        lngTemp0 = lngTemp0 Xor T2(0, (c0 And &HFF&))
        lngTemp1 = lngTemp1 Xor T2(1, (c1 And &HFF&))
        
        .Rotate_64 c0, c1, -56                           ' Rotate right 56 positions
        lngTemp0 = lngTemp0 Xor T1(0, (c0 And &HFF&))
        lngTemp1 = lngTemp1 Xor T1(1, (c1 And &HFF&))
        
        Add64 b0, b1, lngTemp0, lngTemp1                 ' Update b0 & b1 using addition
    End With
    
    ' if((mul)==5) \
    '    {mul5(b0, b1);} \
    ' else \
    '    if((mul)==7) \
    '       {mul7(b0, b1);} \
    '    else \
    '       {mul9(b0, b1)};
    Select Case lngMultiplier
           Case 5: Mul5 b0, b1
           Case 7: Mul7 b0, b1
           Case 9: Mul9 b0, b1
    End Select

End Sub

' #define key_schedule \
Private Sub KeySchedule(ByRef alngMsg() As Long)

    ' key assignments (my reference)
    ' ex:  0, 1   0 - index pointer 0 (left half of number)
    '             1 - index pointer 1 (right half of number)
    '
    '  X00, x01 = 0, 1
    '  X10, x11 = 2, 3
    '  X20, x21 = 4, 5
    '  X30, x31 = 6, 7
    '  X40, x41 = 8, 9
    '  X50, x51 = 10, 11
    '  X60, x61 = 12, 13
    '  X70, x71 = 14, 15

    ' Called by Transform()
    
    Dim lngIdx   As Long
    Dim lngLeft  As Long
    Dim lngRight As Long
    Dim lngLoop  As Long
    Dim lngLoops As Long

    lngIdx = 0   ' Initialize index pointer

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Sub
    End If

    Select Case mlngHashMethod
           Case eTiger3_128, eTiger3_160, eTiger3_192:              lngLoops = 1
           Case eTiger3_224, eTiger3_256, eTiger3_384, eTiger3_512: lngLoops = 2
    End Select
    
    With mobjMath32
        For lngLoop = 1 To lngLoops
    
            ' sub64(x00, x01, x70^0xA5A5A5A5, x71^0xA5A5A5A5); \
            lngLeft = alngMsg(14 + lngIdx)
            lngRight = alngMsg(15 + lngIdx)
            Xor64 lngLeft, lngRight, &HA5A5A5A5, &HA5A5A5A5
            Sub64 alngMsg(0 + lngIdx), alngMsg(1 + lngIdx), lngLeft, lngRight
    
            ' xor64(x10, x11, x00, x01); \
            Xor64 alngMsg(2 + lngIdx), alngMsg(3 + lngIdx), alngMsg(0 + lngIdx), alngMsg(1 + lngIdx)
    
            ' add64(x20, x21, x10, x11); \
            Add64 alngMsg(4 + lngIdx), alngMsg(5 + lngIdx), alngMsg(2 + lngIdx), alngMsg(3 + lngIdx)
    
            ' sub64(x30, x31, x20^((~x10)<<19), ~x21^(((x11)<<19)|((x10)>>13))); \
            lngLeft = alngMsg(4 + lngIdx) Xor .w32Shift((Not (alngMsg(2 + lngIdx))), 19)
            lngRight = (Not (alngMsg(5 + lngIdx)) Xor _
                       (.w32Shift(alngMsg(3 + lngIdx), 19) Or _
                       (.w32Shift(alngMsg(2 + lngIdx), -13))))
            Sub64 alngMsg(6 + lngIdx), alngMsg(7 + lngIdx), lngLeft, lngRight
            
            ' xor64(x40, x41, x30, x31); \
            Xor64 alngMsg(8 + lngIdx), alngMsg(9 + lngIdx), alngMsg(6 + lngIdx), alngMsg(7 + lngIdx)
    
            ' add64(x50, x51, x40, x41); \
            Add64 alngMsg(10 + lngIdx), alngMsg(11 + lngIdx), alngMsg(8 + lngIdx), alngMsg(9 + lngIdx)
    
            ' sub64(x60, x61, ~x50^(((x40)>>23)|((x41)<<9)), x51^((~x41)>>23)); \
            lngLeft = (Not (alngMsg(10 + lngIdx)) Xor _
                      (.w32Shift(alngMsg(8 + lngIdx), -23) Or _
                      (.w32Shift(alngMsg(9 + lngIdx), 9))))
            lngRight = alngMsg(11 + lngIdx) Xor (.w32Shift((Not (alngMsg(9 + lngIdx))), -23))
            Sub64 alngMsg(12 + lngIdx), alngMsg(13 + lngIdx), lngLeft, lngRight
    
            ' xor64(x70, x71, x60, x61); \
            Xor64 alngMsg(14 + lngIdx), alngMsg(15 + lngIdx), alngMsg(12 + lngIdx), alngMsg(13 + lngIdx)
    
            ' add64(x00, x01, x70, x71); \
            Add64 alngMsg(0 + lngIdx), alngMsg(1 + lngIdx), alngMsg(14 + lngIdx), alngMsg(15 + lngIdx)
    
            ' sub64(x10, x11, x00^((~x70)<<19), ~x01^(((x71)<<19)|((x70)>>13))); \
            lngLeft = alngMsg(0 + lngIdx) Xor (.w32Shift((Not (alngMsg(14 + lngIdx))), 19))
            lngRight = (Not (alngMsg(1 + lngIdx) Xor _
                       (.w32Shift(alngMsg(15 + lngIdx), 19)) Or _
                       (.w32Shift(alngMsg(14 + lngIdx), -13))))
            Sub64 alngMsg(2 + lngIdx), alngMsg(3 + lngIdx), lngLeft, lngRight
    
            ' xor64(x20, x21, x10, x11); \
            Xor64 alngMsg(4 + lngIdx), alngMsg(5 + lngIdx), alngMsg(2 + lngIdx), alngMsg(3 + lngIdx)
    
            ' add64(x30, x31, x20, x21); \
            Add64 alngMsg(6 + lngIdx), alngMsg(7 + lngIdx), alngMsg(4 + lngIdx), alngMsg(5 + lngIdx)
    
            ' sub64(x40, x41, ~x30^(((x20)>>23)|((x21)<<9)), x31^((~x21)>>23)); \
            lngLeft = (Not (alngMsg(6 + lngIdx)) Xor _
                      (.w32Shift(alngMsg(4 + lngIdx), -23) Or _
                      (.w32Shift(alngMsg(5 + lngIdx), 9))))
            lngRight = alngMsg(7 + lngIdx) Xor (.w32Shift((Not (alngMsg(5 + lngIdx))), -23))
            Sub64 alngMsg(8 + lngIdx), alngMsg(9 + lngIdx), lngLeft, lngRight
    
            ' xor64(x50, x51, x40, x41); \
            Xor64 alngMsg(10 + lngIdx), alngMsg(11 + lngIdx), alngMsg(8 + lngIdx), alngMsg(9 + lngIdx)
           
            ' add64(x60, x61, x50, x51); \
            Add64 alngMsg(12 + lngIdx), alngMsg(13 + lngIdx), alngMsg(10 + lngIdx), alngMsg(11 + lngIdx)
    
            ' sub64(x70, x71, x60^0x89ABCDEF, x61^0x01234567);
            lngLeft = alngMsg(12 + lngIdx)
            lngRight = alngMsg(13 + lngIdx)
            Xor64 lngLeft, lngRight, &H89ABCDEF, &H1234567
            Sub64 alngMsg(14 + lngIdx), alngMsg(15 + lngIdx), lngLeft, lngRight
    
            lngIdx = lngIdx + 16   ' Increment index pointer
    
            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit For   ' exit FOR..NEXT loop
            End If
    
        Next lngLoop
    End With
    
End Sub

' #define mul5(s0, s1) \
' Multiply by 5
Private Sub Mul5(ByRef s0 As Long, _
                 ByRef s1 As Long)

    ' Called by CalcRound()
    
    Dim lngTemp0 As Long
    Dim lngTemp1 As Long

    ' tempt0 = s0<<2; \
    ' tempt1 = (s1<<2)|(s0>>30); \
    With mobjMath32
        lngTemp0 = .w32Shift(s0, 2)
        lngTemp1 = (.w32Shift(s1, 2)) Or (.w32Shift(s0, -30))
    End With

    ' add64(s0, s1, tempt0, tempt1);
    Add64 s0, s1, lngTemp0, lngTemp1

End Sub

' #define mul7(s0, s1) \
' Multiply by 7
Private Sub Mul7(ByRef s0 As Long, _
                 ByRef s1 As Long)

    ' Called by CalcRound()
    
    Dim lngTemp0 As Long
    Dim lngTemp1 As Long

    ' tempt0 = s0<<3; \
    ' tempt1 = (s1<<3)|(s0>>29); \
    With mobjMath32
        lngTemp0 = .w32Shift(s0, 3)
        lngTemp1 = (.w32Shift(s1, 3)) Or (.w32Shift(s0, -29))
    End With

    ' sub64(tempt0, tempt1, s0, s1); \
    Sub64 lngTemp0, lngTemp1, s0, s1

    s0 = lngTemp0   ' s0 = tempt0; \
    s1 = lngTemp1   ' s1 = tempt1;

End Sub

' #define mul9(s0, s1) \
' Multiply by 9
Private Sub Mul9(ByRef s0 As Long, _
                 ByRef s1 As Long)

    ' Called by CalcRound()
    
    Dim lngTemp0 As Long
    Dim lngTemp1 As Long

    ' tempt0 = s0<<3; \
    ' tempt1 = (s1<<3)|(s0>>29); \
    With mobjMath32
        lngTemp0 = .w32Shift(s0, 3)
        lngTemp1 = (.w32Shift(s1, 3)) Or (.w32Shift(s0, -29))
    End With

    ' add64(s0, s1, tempt0, tempt1);
    Add64 s0, s1, lngTemp0, lngTemp1

End Sub

' #define add64(s0, s1, p0, p1) \
' s0,s1 values are manipulated and returned
Private Sub Add64(ByRef s0 As Long, _
                  ByRef s1 As Long, _
                  ByVal p0 As Long, _
                  ByVal p1 As Long)

    ' Called by CalcRound()
    '           FeedForward()
    '           KeySchedule()
    '           Mul5()
    '           Mul9()
    
    Dim lngTemp  As Long
    Dim lngCarry As Long

    With mobjMath32
        lngTemp = p0                     ' temps0 = (p0); \
        s0 = .UnsignedAdd(s0, lngTemp)   ' s0 += temps0; \

        ' tcarry = s0 < temps0; \
        If s0 < lngTemp Then
            lngCarry = s0
        Else
            lngCarry = lngTemp
        End If

        ' s1+=(p1) + tcarry;
        lngTemp = .UnsignedAdd(p1, lngCarry)
        s1 = .UnsignedAdd(s1, lngTemp)
    End With

End Sub

' #define sub64(s0, s1, p0, p1) \
' s0,s1 values are manipulated and returned
Private Sub Sub64(ByRef s0 As Long, _
                  ByRef s1 As Long, _
                  ByVal p0 As Long, _
                  ByVal p1 As Long)

    ' Called by CalcRound()
    '           FeedForward()
    '           KeySchedule()
    '           Mul7()
    
    Dim lngTemp  As Long
    Dim lngCarry As Long

    lngTemp = p0  ' temps0 = (p0); \

    ' tcarry = s0 < temps0; \
    If s0 < lngTemp Then
        lngCarry = s0
    Else
        lngCarry = lngTemp
    End If

    ' s0 -= temps0; \
    ' s1 -= (p1) + tcarry;
    With mobjMath32
        s0 = .UnsignedSubtract(s0, lngTemp)
        lngTemp = .UnsignedAdd(p1, lngCarry)
        s1 = .UnsignedSubtract(s1, lngTemp)
    End With

End Sub

' #define xor64(s0, s1, p0, p1) \
' s0,s1 values are manipulated and returned
Private Sub Xor64(ByRef s0 As Long, _
                  ByRef s1 As Long, _
                  ByVal p0 As Long, _
                  ByVal p1 As Long)

    ' Called by CalcRound()
    '           FeedForward()
    '           KeySchedule()
    
    s0 = s0 Xor p0   ' s0 ^= (p0); \
    s1 = s1 Xor p1   ' s1 ^= (p1);

End Sub

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the data to be processed.
'
' Parameters:    curAmtLeft - Amount of data left
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Called by HashFile()
    '           HashString()
    
    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= MB_5: GetBlockSize = MB_5
           Case Else:       GetBlockSize = CLng(curAmtLeft)
    End Select

End Function

' ***************************************************************************
' Routine:       LoadWorkArrays
'
' Description:   Preload the Work array and Constants array for 32-bit 
'                processing with predetermined values to be used in the 
'                hashing calculations. 
' 
'                Got this explanation from Ask Dr. Math web site on how to
'                calculate the hex representation of the fractional parts of
'                a whole number.
'                http://mathforum.org/dr.math/ 
' 
'                There is a simple process for converting a base 10 decimal
'                fraction to a "decimal" fraction in another base.  In this
'                case, hex (16) will be our new base.  We repeatedly multiply
'                our decimal fraction by the new base, picking off the whole
'                number part each time as the next digit of the final output. 
'
' ---------------------------------------------------------------------------- 
'            My idea for data to fill the work arrays is to calculate the
'            hex representation of the fractional part of the CUBE ROOTS
'            of the 512th thru 589th prime numbers.
'
' Example:   512th Prime number = 3671
'            CUBE ROOT of 3671 = 15.426288936629509912318786195
'                                 |
'                                 +-- Whole number is dropped
'
'            0.426288936629509912318786195 * 16 =  6.82062298607215859710057912
'            0.82062298607215859710057912  * 16 = 13.12996777715453755360926592
'            0.12996777715453755360926592  * 16 =  2.07948443447260085774825472
'            0.07948443447260085774825472  * 16 =  1.27175095156161372397207552
'            0.27175095156161372397207552  * 16 =  4.34801522498581958355320832
'            0.34801522498581958355320832  * 16 =  5.56824359977311333685133312
'            0.56824359977311333685133312  * 16 =  9.09189759636981338962132992
'            0.09189759636981338962132992  * 16 =  1.47036154191701423394127872
'            0.47036154191701423394127872  * 16 =  7.52578467067222774306045952
'            0.52578467067222774306045952  * 16 =  8.41255473075564388896735232
'            0.41255473075564388896735232  * 16 =  6.60087569209030222347763712
'            0.60087569209030222347763712  * 16 =  9.61401107344483557564219392
'            0.61401107344483557564219392  * 16 =  9.82417717511736921027510272
'            0.82417717511736921027510272  * 16 = 13.18683480187790736440164352
'            0.18683480187790736440164352  * 16 =  2.98935683004651783042629632
'            0.98935683004651783042629632  * 16 = 15.82970928074428528682074112
'                                                  |
'             Whole numbers are converted to hex --+
'
'            The hex representation of the fractional parts
'            of the CUBE ROOT of 3671:  6d21459178699d2f
'                                                |      |
'                                                --------
'                     Last 32-bits (8 characters) --+
'
'            Ex:  malngHash(0) = &H78699D2F
' 
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Combined all Tiger hashes into a single module.
'              - Kept original data for the constant arrays.
' 05-Sep-2016  Kenneth Ives  kenaso@tx.rr.com
'              Created unique initialization values for each of
'              the Tiger3 hashes.
' ***************************************************************************
Private Sub LoadWorkArrays()

    Dim lngIdx   As Long
    Dim lngIndex As Long
    Dim strData  As String
    Dim avntData As Variant

    Select Case mlngHashMethod
           Case eTiger3_128
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 512th through 517th prime numbers.
                ReDim malngHash(6)
                malngHash(0) = &H78699D2F
                malngHash(1) = &HAFBA5E23
                malngHash(2) = &H6950D47A
                malngHash(3) = &H935121DF
                malngHash(4) = &H4A5FCDEB
                malngHash(5) = &H2D4A6873

           Case eTiger3_160
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 518th through 523rd prime numbers.
                ReDim malngHash(6)
                malngHash(0) = &H45DE1EEA
                malngHash(1) = &HF5358C61
                malngHash(2) = &H9DA35205
                malngHash(3) = &H5A116F23
                malngHash(4) = &H12182662
                malngHash(5) = &HB716845C

           Case eTiger3_192
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 524th through 529rd prime numbers.
                ReDim malngHash(6)
                malngHash(0) = &HADED7439
                malngHash(1) = &HDD462C71
                malngHash(2) = &HEF6339E9
                malngHash(3) = &HC869BE18
                malngHash(4) = &H53F7FA88
                malngHash(5) = &HB54501AE

           Case eTiger3_224
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 530th through 541st prime numbers.
                ReDim malngHash(12)
                malngHash(0) = &HB54501AE
                malngHash(1) = &HC2B3B044
                malngHash(2) = &H6116716D
                malngHash(3) = &H99B2C5AC
                malngHash(4) = &HCBB5EC26
                malngHash(5) = &H88E13CBC
                malngHash(6) = &H1B77BA24
                malngHash(7) = &H415A1AD3
                malngHash(8) = &HE33A5BA3
                malngHash(9) = &HCA8AABCE
                malngHash(10) = &H3D5A464D
                malngHash(11) = &H2E64F2E

           Case eTiger3_256
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 542nd through 553rd prime numbers.
                ReDim malngHash(12)
                malngHash(0) = &H4DE7C5AF
                malngHash(1) = &H27B0150A
                malngHash(2) = &H1BCCD63
                malngHash(3) = &HF24C9F7
                malngHash(4) = &H6CB4E92C
                malngHash(5) = &H22E67181
                malngHash(6) = &H3DAF9FC2
                malngHash(7) = &H3BB50EC3
                malngHash(8) = &H9533E090
                malngHash(9) = &HAB3FA038
                malngHash(10) = &H7F41DB2
                malngHash(11) = &H856E3A8

           Case eTiger3_384
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 554th through 571st prime numbers.
                ReDim malngHash(18)
                malngHash(0) = &H629EC872
                malngHash(1) = &H6270F054
                malngHash(2) = &H1C9392CE
                malngHash(3) = &H14ADCCA1
                malngHash(4) = &H9E9F11E1
                malngHash(5) = &H776B865A
                malngHash(6) = &HED449FEC
                malngHash(7) = &HAED905FA
                malngHash(8) = &H1BE27B8C
                malngHash(9) = &HD54DFC34
                malngHash(10) = &H551FF8A9
                malngHash(11) = &HAA755CA9
                malngHash(12) = &HD3794395
                malngHash(13) = &H47600C5A
                malngHash(14) = &H7D08B7E2
                malngHash(15) = &H3EB1B6CE
                malngHash(16) = &HA4675BCD
                malngHash(17) = &HAE901657

           Case eTiger3_512
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the CUBE ROOTS
                ' of the 572nd through 589th prime numbers.
                ReDim malngHash(18)
                malngHash(0) = &H748B3C4B
                malngHash(1) = &H7EBA1D39
                malngHash(2) = &HB06DDF92
                malngHash(3) = &HA168C4C1
                malngHash(4) = &H40BEF5FD
                malngHash(5) = &H2171CD6E
                malngHash(6) = &H35918AE4
                malngHash(7) = &HB72ED706
                malngHash(8) = &H8BD85572
                malngHash(9) = &HE7A9DF4D
                malngHash(10) = &H6167D94D
                malngHash(11) = &H35ADD8CA
                malngHash(12) = &HCC396EFB
                malngHash(13) = &H2E113BF0
                malngHash(14) = &HFE386A17
                malngHash(15) = &H8B9CBCDF
                malngHash(16) = &HEC9E9646
                malngHash(17) = &H2D545F3A
    End Select

    ' See if arrays are already loaded
    If CBool(IsArrayInitialized(T1)) Then
        If T1(0, 0) > 0 Then
            Exit Sub
        End If
    End If
    
    ' Size constant arrays
    ReDim T1(0 To 1, MAX_ARRAY)
    ReDim T2(0 To 1, MAX_ARRAY)
    ReDim T3(0 To 1, MAX_ARRAY)
    ReDim T4(0 To 1, MAX_ARRAY)

    strData = vbNullString   ' Verify string is empty
    avntData = Empty         ' Verify variant is empty
    lngIdx = 0

    ' Original data for constant arrays
    ' Load T1 table array
    strData = strData & "02AAB17C F7E90C5E AC424B03 E243A8EC 72CD5BE3 0DD5FCD3 6D019B93 F6F97F3A "
    strData = strData & "CD9978FF D21F9193 7573A1C9 708029E2 B164326B 922A83C3 46883EEE 04915870 "
    strData = strData & "EAACE305 7103ECE6 C54169B8 08A3535C 4CE75491 8DDEC47C 0AA2F4DF DC0DF40C "
    strData = strData & "10B76F18 A74DBEFA C6CCB623 5AD1AB6A 13726121 572FE2FF 1A488C6F 199D921E "
    strData = strData & "4BC9F9F4 DA0007CA 26F5E6F6 E85241C7 859079DB EA5947B6 4F1885C5 C99E8C92 "
    strData = strData & "D78E761E A96F864B 8E36428C 52B5C17D 69CF6827 373063C1 B607C93D 9BB4C56E "
    strData = strData & "7D820E76 0E76B5EA 645C9CC6 F07FDC42 BF38A078 243342E0 5F6B343C 9D2E7D04 "
    strData = strData & "F2C28AEB 600B0EC6 6C0ED85F 7254BCAC 71592281 A4DB4FE5 1967FA69 CE0FED9F "
    strData = strData & "FD5293F8 B96545DB C879E9D7 F2A7600B 86024892 0193194E A4F9533B 2D9CC0B3 "
    strData = strData & "9053836C 15957613 DB6DCF8A FC357BF1 18BEEA7A 7A370F57 037117CA 50B99066 "
    strData = strData & "6AB30A97 74424A35 F4E92F02 E325249B 7739DB07 061CCAE1 D8F3B49C ECA42A05 "
    strData = strData & "BD56BE3F 51382F73 45FAED58 43B0BB28 1C813D5C 11BF1F83 8AF0E4B6 D75FA169 "
    strData = strData & "33EE18A4 87AD9999 3C26E8EA B1C94410 B510102B C0A822F9 141EEF31 0CE6123B "
    strData = strData & "FC65B900 59DDB154 E0158640 C5E0E607 884E0798 26C3A3CF 930D0D95 23C535FD "
    strData = strData & "35638D75 4E9A2B00 4085FCCF 40469DD5 C4B17AD2 8BE23A4C CAB2F0FC 6A3E6A2E "
    strData = strData & "2860971A 6B943FCD 3DDE6EE2 12E30446 6222F32A E01765AE 5D550BB5 478308FE "
    strData = strData & "A9EFA98D A0EDA22A C351A716 86C40DA7 1105586D 9C867C84 DCFFEE85 FDA22853 "
    strData = strData & "CCFBD026 2C5EEF76 BAF294CB 8990D201 E69464F5 2AFAD975 94B013AF DF133E14 "
    strData = strData & "06A7D1A3 2823C958 6F95FE51 30F61119 D92AB34E 462C06C0 ED7BDE33 887C71D2 "
    strData = strData & "79746D6E 6518393E 5BA41938 5D713329 7C1BA6B9 48A97564 31987C19 7BFDAC67 "
    strData = strData & "DE6C23C4 4B053D02 581C49FE D002D64D DD474D63 38261571 AA4546C3 E473D062 "
    strData = strData & "928FCE34 9455F860 48161BBA CAAB94D9 63912430 770E6F68 6EC8A5E6 02C6641C "
    strData = strData & "87282515 337DDD2B 2CDA6B42 034B701B B03D37C1 81CB096D E1084382 66C71C6F "
    strData = strData & "2B3180C7 EB51B255 DF92B82F 96C08BBC 5C68C8C0 A632F3BA 5504CC86 1C3D0556 "
    strData = strData & "ABBFA4E5 5FB26B8F 41848B0A B3BACEB4 B334A273 AA445D32 BCA696F0 A85AD881 "
    strData = strData & "24F6EC65 B528D56C 0CE1512E 90F4524A 4E9DD79D 5506D35A 258905FA C6CE9779 "
    strData = strData & "2019295B 3E109B33 F8A9478B 73A054CC 2924F2F9 34417EB0 3993357D 536D1BC4 "
    strData = strData & "38A81AC2 1DB6FF8B 47C4FBF1 7D6016BF 1E0FAADD 7667E3F5 7ABCFF62 938BEB96 "
    strData = strData & "A78DAD94 8FC179C9 8F1F98B7 2911E50D 61E48EAE 27121A91 4D62F7AD 31859808 "
    strData = strData & "ECEBA345 EF5CEAEB F5CEB25E BC9684CE F633E20C B7F76221 A32CDF06 AB8293E4 "
    strData = strData & "985A202C A5EE2CA4 CF0B8447 CC8A8FB1 9F765244 979859A3 A8D516B1 A1240017 "
    strData = strData & "0BD7BA3E BB5DC726 E54BCA55 B86ADB39 1D7A3AFD 6C478063 519EC608 E7669EDD "
    strData = strData & "0E5715A2 D149AA23 177D4571 848FF194 EEB55F32 41014C22 0F5E5CA1 3A6E2EC2 "
    strData = strData & "8029927B 75F5C361 AD139FAB C3D6E436 0D5DF1A9 4CCF402F 3E8BD948 BEA5DFC8 "
    strData = strData & "A5A0D357 BD3FF77E A2D12E25 1F74F645 66FD9E52 5E81A082 2E0C90CE 7F687A49 "
    strData = strData & "C2E8BCBE BA973BC5 000001BC E509745F 423777BB E6DAB3D6 D1661C7E AEF06EB5 "
    strData = strData & "A1781F35 4DAACFD8 2D11284A 2B16AFFC F1FC4F67 FA891D1F 73ECC25D CB920ADA "
    strData = strData & "AE610C22 C2A12651 96E0A810 D356B78A 5A9A381F 2FE7870F D5AD62ED E94E5530 "
    strData = strData & "D225E5E8 368D1427 65977B70 C7AF4631 99F889B2 DE39D74F 233F30BF 54E1D143 "
    strData = strData & "9A9675D3 D9A63C97 5470554F F334F9A8 166ACB74 4A4F5688 70C74CAA B2E4AEAD "
    strData = strData & "F0D09164 6F294D12 57B82A89 684031D1 EFD95A5A 61BE0B6B 2FBD12E9 69F2F29A "
    strData = strData & "9BD37013 FEFF9FE8 3F9B0404 D6085A06 4940C1F3 166CFE15 09542C4D CDF3DEFB "
    strData = strData & "B4C52183 85CD5CE3 C935B7DC 4462A641 3417F8A6 8ED3B63F B8095929 5B215B40 "
    strData = strData & "F99CDAEF 3B8C8572 018C0614 F8FCB95D 1B14ACCD 1A3ACDF3 84D471F2 00BB732D "
    strData = strData & "C1A3110E 95E8DA16 430A7220 BF1A82B8 B77E090D 39DF210E 5EF4BD9F 3CD05E9D "
    strData = strData & "9D4FF6DA 7E57A444 DA1D60E1 83D4A5F8 B287C384 17998E47 FE3EDC12 1BB31886 "
    strData = strData & "C7FE3CCC 980CCBEF E46FB590 189BFD03 3732FD46 9A4C57DC 7EF700A0 7CF1AD65 "
    strData = strData & "59C64468 A31D8859 762FB0B4 D45B61F6 155BAED0 99047718 68755E4C 3D50BAA6 "
    strData = strData & "E9214E7F 22D8B4DF 2ADDBF53 2EAC95F4 32AE3909 B4BD0109 834DF537 B08E3450 "
    strData = strData & "FA209DA8 4220728D 9E691D9B 9EFE23F7 0446D288 C4AE8D7F 7B4CC524 E169785B "
    strData = strData & "21D87F01 35CA1385 CEBB400F 137B8AA5 272E2B66 580796BE 36122641 25C2B0DE "
    strData = strData & "057702BD AD1EFBB2 D4BABB8E ACF84BE9 91583139 641BC67B 8BDC2DE0 8036E024 "
    strData = strData & "603C8156 F49F68ED F7D236F7 DBEF5111 9727C459 8AD21E80 A08A0896 670A5FD7 "
    strData = strData & "CB4A8F43 09EBA9CB 81AF564B 0F7036A1 C0B99AA7 78199ABD 959F1EC8 3FC8E952 "
    strData = strData & "8C505077 794A81B9 3ACAAF8F 056338F0 07B43F50 627A6778 4A44AB49 F5ECCC77 "
    strData = strData & "3BC3D6E4 B679EE98 9CC0D4D1 CF14108C 4406C00B 206BC8A0 82A18854 C8D72D89 "
    strData = strData & "67E366B3 5C3C432C B923DD61 102B37F2 56AB2779 D884271D BE83E1B0 FF1525AF "
    strData = strData & "FB7C65D4 217E49A9 6BDBE0E7 6D48E7D4 08DF8287 45D9179E 22EA6A9A DD53BD34 "
    strData = strData & "E36E141C 5622200A 7F805D1B 8CB750EE AFE5C7A5 9F58E837 E27F996A 4FB1C23C "
    strData = strData & "D3867DFB 0775F0D0 D0E673DE 6E88891A 123AEB9E AFB86C25 30F1D5D5 C145B895 "
    strData = strData & "BB434A2D EE7269E7 78CB67EC F931FA38 F33B0372 323BBF9C 52D66336 FB279C74 "
    strData = strData & "505F33AC 0AFB4EAA E8A5CD99 A2CCE187 53497480 1E2D30BB 8D2D5711 D5876D90 "
    strData = strData & "1F1A4128 91BC038E D6E2E71D 82E56648 74036C3A 497732B7 89B67ED9 6361F5AB "
    strData = strData & "FFED95D8 F1EA02A2 E72B3BD6 1464D43D A6300F17 0BDC4820 EBC18760 ED78A77A "

    avntData = Split(strData, Chr$(32))   ' Load variant array

    ' load T1 table array
    For lngIndex = 0 To 511 Step 2

        ' Convert variant data to Long Integer
        T1(0, lngIdx) = CLng("&H" & avntData(lngIndex))
        T1(1, lngIdx) = CLng("&H" & avntData(lngIndex + 1))
        lngIdx = lngIdx + 1

    Next lngIndex

    strData = vbNullString   ' Verify string is empty
    avntData = Empty         ' Verify variant is empty
    lngIdx = 0

    ' Load T2 table array
    strData = strData & "E6A6BE5A 05A12138 B5A122A5 B4F87C98 563C6089 140B6990 4C46CB2E 391F5DD5 "
    strData = strData & "D932ADDB C9B79434 08EA70E4 2015AFF5 D765A667 3E478CF1 C4FB757E AB278D99 "
    strData = strData & "DF11C686 2D6E0692 DDEB84F1 0D7F3B16 6F2EF604 A665EA04 4A8E0F0F F0E0DFB3 "
    strData = strData & "A5EDEEF8 3DBCBA51 FC4F0A2A 0EA4371E E83E1DA8 5CB38429 DC8FF882 BA1B1CE2 "
    strData = strData & "CD45505E 8353E80D 18D19A00 D4DB0717 34A0CFED A5F38101 0BE77E51 8887CAF2 "
    strData = strData & "1E341438 B3C45136 E05797F4 9089CCF9 FFD23F9D F2591D14 543DDA22 8595C5CD "
    strData = strData & "661F81FD 99052A33 8736E641 DB0F7B76 15227725 418E5307 E25F7F46 162EB2FA "
    strData = strData & "48A8B212 6C13D9FE AFDC5417 92E76EEA 03D912BF C6D1898F 31B1AAFA 1B83F51B "
    strData = strData & "F1AC2796 E42AB7D9 40A3A7D7 FCD2EBAC 1056136D 0AFBBCC5 7889E1DD 9A6D0C85 "
    strData = strData & "D3352578 2A7974AA A7E25D09 078AC09B BD4138B3 EAC6EDD0 920ABFBE 71EB9E70 "
    strData = strData & "A2A5D0F5 4FC2625C C054E36B 0B1290A3 F6DD59FF 62FE932B 35373545 11A8AC7D "
    strData = strData & "CA845E91 72FADCD4 84F82B60 329D20DC 79C62CE1 CD672F18 8B09A2AD D124642C "
    strData = strData & "D0C1E96A 19D9E726 5A786A9B 4BA9500C 0E020336 634C43F3 C17B474A EB66D822 "
    strData = strData & "6A731AE3 EC9BAAC2 8226667A E0840258 67D45676 91CAECA5 1D94155C 4875ADB5 "
    strData = strData & "6D00FD98 5B813FDF 51286EFC B774CD06 5E883447 1FA744AF F72CA0AE E761AE2E "
    strData = strData & "BE40E4CD AEE8E09A E9970BBB 5118F665 726E4BEB 33DF1964 703B0007 29199762 "
    strData = strData & "4631D816 F5EF30A7 B880B5B5 1504A6BE 641793C3 7ED84B6C 7B21ED77 F6E97D96 "
    strData = strData & "77630631 2EF96B73 AE528948 E86FF3F4 53DBD7F2 86A3F8F8 16CADCE7 4CFC1063 "
    strData = strData & "005C19BD FA52C6DD 68868F5D 64D46AD3 3A9D512C CF1E186A 367E62C2 385660AE "
    strData = strData & "E359E7EA 77DCB1D7 526C0773 749ABE6E 735AE5F9 D09F734B 493FC7CC 8A558BA8 "
    strData = strData & "B0B9C153 3041AB45 321958BA 470A59BD 852DB00B 5F46C393 91209B2B D336B0E5 "
    strData = strData & "6E604F7D 659EF19F B99A8AE2 782CCB24 CCF52AB6 C814C4C7 4727D9AF BE11727B "
    strData = strData & "7E950D0C 0121B34D 756F4356 70AD471F F5ADD442 615A6849 4E87E099 80B9957A "
    strData = strData & "2ACFA1DF 50AEE355 D898263A FD2FD556 C8F4924D D80C8FD6 CF99CA3D 754A173A "
    strData = strData & "FE477BAC AF91BF3C ED5371F6 D690C12D 831A5C28 5E687094 C5D3C90A 3708A0A4 "
    strData = strData & "0F7F9037 17D06580 19F9BB13 B8FDF27F B1BD6F1B 4D502843 1C761BA3 8FFF4012 "
    strData = strData & "0D1530C4 E2E21F3B 8943CE69 A7372C8A E5184E11 FEB5CE66 618BDB80 BD736621 "
    strData = strData & "7D29BAD6 8B574D0B 81BB613E 25E6FE5B 071C9C10 BC07913F C7BEEB79 09AC2D97 "
    strData = strData & "C3E58D35 3BC5D757 EB017892 F38F61E8 D4EFFB9C 9B1CC21A 99727D26 F494F7AB "
    strData = strData & "A3E063A2 956B3E03 9D4A8B9A 4AA09C30 3F6AB7D5 00090FB4 9CC0F2A0 57268AC0 "
    strData = strData & "3DEE9D2D EDBF42D1 330F49C8 7960A972 C6B27202 87421B41 0AC59EC0 7C00369C "
    strData = strData & "EF4EAC49 CB353425 F450244E EF0129D8 8ACC46E5 CAF4DEB6 2FFEAB63 989263F7 "
    strData = strData & "8F7CB9FE 5D7A4578 5BD8F764 4E634635 427A7315 BF2DC900 17D0C4AA 2125261C "
    strData = strData & "3992486C 93518E50 B4CBFEE0 A2D7D4C3 7C75D620 2C5DDD8D DBC295D8 E35B6C61 "
    strData = strData & "60B369D3 02032B19 CE42685F DCE44132 06F3DDB9 DDF65610 8EA4D21D B5E148F0 "
    strData = strData & "20B0FCE6 2FCD496F 2C1B9123 58B0EE31 B28317B8 18F5A308 A89C1E18 9CA6D2CF "
    strData = strData & "0C6B1857 6AAADBC8 B65DEAA9 1299FAE3 FB2B794B 7F1027E7 04E4317F 443B5BEB "
    strData = strData & "4B852D32 5939D0A6 D5AE6BEE FB207FFC 309682B2 81C7D374 BAE309A1 94C3B475 "
    strData = strData & "8CC3F97B 13B49F05 98A9422F F8293967 244B16B0 1076FF7C F8BF571C 663D67EE "
    strData = strData & "1F0D6758 EEE30DA1 C9B611D9 7ADEB9B7 B7AFD588 7B6C57A2 6290AE84 6B984FE1 "
    strData = strData & "94DF4CDE ACC1A5FD 058A5BD1 C5483AFF 63166CC1 42BA3C37 8DB8526E B2F76F40 "
    strData = strData & "E1088003 6F0D6D4E 9E0523C9 971D311D 45EC2824 CC7CD691 575B8359 E62382C9 "
    strData = strData & "FA9E400D C4889995 D1823ECB 45721568 DAFD983B 8206082F AA7D2908 2386A8CB "
    strData = strData & "269FCD44 03B87588 1B91F5F7 28BDD1E0 E4669F39 040201F6 7A1D7C21 8CF04ADE "
    strData = strData & "65623C29 D79CE5CE 23684490 96C00BB1 AB9BF187 9DA503BA BC23ECB1 A458058E "
    strData = strData & "9A58DF01 BB401ECC A070E868 A85F143D 4FF18830 7DF2239E 14D565B4 1A641183 "
    strData = strData & "EE133374 52701602 950E3DCF 3F285E09 59930254 B9C80953 3BF29940 8930DA6D "
    strData = strData & "A955943F 53691387 A15EDECA A9CB8784 29142127 352BE9A0 76F0371F FF4E7AFB "
    strData = strData & "0239F450 274F2228 BB073AF0 1D5E868B BFC80571 C10E96C1 D2670885 68222E23 "
    strData = strData & "9671A3D4 8E80B5B0 55B5D38A E193BB81 693AE2D0 A18B04B8 5C48B4EC ADD5335F "
    strData = strData & "FD743B19 4916A1CA 25770181 34BE98C4 E77987E8 3C54A4AD 28E11014 DA33E1B9 "
    strData = strData & "270CC59E 226AA213 71495F75 6D1A5F60 9BE853FB 60AFEF77 ADC786A7 F7443DBF "
    strData = strData & "09044561 73B29A82 58BC7A66 C232BD5E F306558C 673AC8B2 41F639C6 B6C9772A "
    strData = strData & "216DEFE9 9FDA35DA 11640CC7 1C7BE615 93C43694 565C5527 EA038E62 46777839 "
    strData = strData & "F9ABF3CE 5A3E2469 741E768D 0FD312D2 0144B883 CED652C6 C20B5A5B A33F8552 "
    strData = strData & "1AE69633 C3435A9D 97A28CA4 088CFDEC 8824A43C 1E96F420 37612FA6 6EEEA746 "
    strData = strData & "6B4CB165 F9CF0E5A 43AA1C06 A0ABFB4A 7F4DC26F F162796B 6CBACC8E 54ED9B0F "
    strData = strData & "A6B7FFEF D2BB253E 2E25BC95 B0A29D4F 86D6A58B DEF1388C DED74AC5 76B6F054 "
    strData = strData & "8030BDBC 2B45805D 3C81AF70 E94D9289 3EFF6DDA 9E3100DB B38DC39F DFCC8847 "
    strData = strData & "12388552 8D17B87E F2DA0ED2 40B1B642 44CEFADC D54BF9A9 1312200E 433C7EE6 "
    strData = strData & "9FFCC84F 3A78C748 F0CD1F72 248576BB EC697405 3638CFE4 2BA7B67C 0CEC4E4C "
    strData = strData & "AC2F4DF3 E5CE32ED CB33D143 26EA4C11 A4E9044C C77E58BC 5F513293 D934FCEF "
    strData = strData & "5DC96455 06E55444 50DE418F 317DE40A 388CB31A 69DDE259 2DB4A834 55820A86 "
    strData = strData & "9010A91E 84711AE9 4DF7F0B7 B1498371 D62A2EAB C0977179 22FAC097 AA8D5C0E "

    avntData = Split(strData, Chr$(32))   ' Load variant array

    ' load T2 table array
    For lngIndex = 0 To 511 Step 2

        ' Convert variant data to Long Integer
        T2(0, lngIdx) = CLng("&H" & avntData(lngIndex))
        T2(1, lngIdx) = CLng("&H" & avntData(lngIndex + 1))
        lngIdx = lngIdx + 1

    Next lngIndex

    strData = vbNullString   ' Verify string is empty
    avntData = Empty         ' Verify variant is empty
    lngIdx = 0

    ' Load T3 table array
    strData = strData & "F49FCC2F F1DAF39B 487FD5C6 6FF29281 E8A30667 FCDCA83F 2C9B4BE3 D2FCCE63 "
    strData = strData & "DA3FF74B 93FBBBC2 2FA165D2 FE70BA66 A103E279 970E93D4 BECDEC77 B0E45E71 "
    strData = strData & "CFB41E72 3985E497 B70AAA02 5EF75017 D42309F0 3840B8E0 8EFC1AD0 35898579 "
    strData = strData & "96C6920B E2B2ABC5 66AF4163 375A9172 2174ABDC CA7127FB B33CCEA6 4A72FF41 "
    strData = strData & "F04A4933 083066A5 8D970ACD D7289AF5 8F96E8E0 31C8C25E F3FEC022 76875D47 "
    strData = strData & "EC7BF310 056190DD F5ADB0AE BB0F1491 9B50F885 0FD58892 49754883 58B74DE8 "
    strData = strData & "A3354FF6 91531C61 0702BBE4 81D2C6EE 89FB2405 7DEDED98 AC307513 8596E902 "
    strData = strData & "1D2D3580 172772ED EB738FC2 8E6BC30D 5854EF8F 63044326 9E5C5232 5ADD3BBE "
    strData = strData & "90AA53CF 325C4623 C1D24D51 349DD067 2051CFEE A69EA624 13220F0A 862E7E4F "
    strData = strData & "CE393994 04E04864 D9C42CA4 7086FCB7 685AD223 8A03E7CC 066484B2 AB2FF1DB "
    strData = strData & "FE9D5D70 EFBF79EC 5B13B9DD 9C481854 15F0D475 ED1509AD 0BEBCD06 0EC79851 "
    strData = strData & "D58C6791 183AB7F8 D1187C50 52F3EEE4 C95D1192 E54E82FF 86EEA14C B9AC6CA2 "
    strData = strData & "3485BEB1 53677D5D DD191D78 1F8C492A F60866BA A784EBF9 518F643B A2D08C74 "
    strData = strData & "8852E956 E1087C22 A768CB8D C410AE8D 38047726 BFEC8E1A A67738B4 CD3B45AA "
    strData = strData & "AD16691C EC0DDE19 C6D43193 80462E07 C5A5876D 0BA61938 16B9FA1F A58FD840 "
    strData = strData & "188AB117 3CA74F18 ABDA2F98 C99C021F 3E0580AB 134AE816 5F3B05B7 73645ABB "
    strData = strData & "2501A2BE 5575F2F6 1B2F7400 4E7E8BA9 1CD75803 71E8D953 7F6ED895 62764E30 "
    strData = strData & "B15926FF 596F003D 9F65293D A8C5D6B9 6ECEF04D D690F84C 4782275F FF33AF88 "
    strData = strData & "E4143308 3F820801 FD0DFE40 9A1AF9B5 4325A334 2CDB396B 8AE77E62 B301B252 "
    strData = strData & "C36F9E9F 6655615A 85455A2D 92D32C09 F2C7DEA9 49477485 63CFB4C1 33A39EBA "
    strData = strData & "83B040CC 6EBC5462 3B9454C8 FDB326B0 56F56A9E 87FFD78C 2DC2940D 99F42BC6 "
    strData = strData & "98F7DF09 6B096E2D 19A6E01E 3AD852BF 42A99CCB DBD4B40B A59998AF 45E9C559 "
    strData = strData & "366295E8 07D93186 6B48181B FAA1F773 1FEC57E2 157A0A1D 4667446A F6201AD5 "
    strData = strData & "E615EBCA CFB0F075 B8F31F4F 68290778 22713ED6 CE22D11E 3057C1A7 2EC3C93B "
    strData = strData & "CB46ACC3 7C3F1F2F DBB893FD 02AAF50E 331FD92E 600B9FCF A498F961 48EA3AD6 "
    strData = strData & "A8D8426E 8B6A83EA A089B274 B7735CDC 87F6B373 1E524A11 118808E5 CBC96749 "
    strData = strData & "9906E4C7 B19BD394 AFED7F7E 9B24A20C 6509EADE EB3644A7 6C1EF1D3 E8EF0EDE "
    strData = strData & "B9C97D43 E9798FB4 A2F2D784 740C28A3 7B849647 6197566F 7A5BE3E6 B65F069D "
    strData = strData & "F96330ED 78BE6F10 EEE60DE7 7A076A15 2B4BEE4A A08B9BD0 6A56A63E C7B8894E "
    strData = strData & "02121359 BA34FEF4 4CBF99F8 283703FC 39807135 0CAF30C8 D0A77A89 F017687A "
    strData = strData & "F1C1A9EB 9E423569 8C797628 2DEE8199 5D1737A5 DD1F7ABD 4F53433C 09A9FA80 "
    strData = strData & "FA8B0C53 DF7CA1D9 3FD9DCBC 886CCB77 C040917C A91B4720 7DD00142 F9D1DCDF "
    strData = strData & "8476FC1D 4F387B58 23F8E7C5 F3316503 032A2244 E7E37339 5C87A5D7 50F5A74B "
    strData = strData & "082B4CC4 3698992E DF917BEC B858F63C 3270B8FC 5BF86DDA 10AE72BB 29B5DD76 "
    strData = strData & "576AC94E 7700362B 1AD112DA C61EFB8F 691BC30E C5FAA427 FF246311 CC327143 "
    strData = strData & "3142368E 30E53206 71380E31 E02CA396 958D5C96 0AAD76F1 F8D6F430 C16DA536 "
    strData = strData & "C8FFD13F 1BE7E1D2 7578AE66 004DDBE1 05833F01 067BE646 BB34B5AD 3BFE586D "
    strData = strData & "095F34C9 A12B97F0 247AB645 25D60CA8 DCDBC6F3 017477D1 4A2E14D4 DECAD24D "
    strData = strData & "BDB5E6D9 BE0A1EEB 2A7E70F7 794301AB DEF42D8A 270540FD 01078EC0 A34C22C1 "
    strData = strData & "E5DE511A F4C16387 7EBB3A52 BD9A330A 77697857 AA7D6435 004E8316 03AE4C32 "
    strData = strData & "E7A21020 AD78E312 9D41A70C 6AB420F2 28E06C18 EA1141E6 D2B28CBD 984F6B28 "
    strData = strData & "26B75F6C 446E9D83 BA47568C 4D418D7F D80BADBF E6183D8E 0E206D7F 5F166044 "
    strData = strData & "E258A439 11CBCA3E 723A1746 B21DC0BC C7CAA854 F5D7CDD3 7CAC3288 3D261D9C "
    strData = strData & "7690C264 23BA942C 17E55524 478042B8 E0BE4776 56A2389F 4D289B5E 67AB2DA0 "
    strData = strData & "44862B9C 8FBBFD31 B47CC804 9D141365 822C1B36 2B91C793 4EB14655 FB13DFD8 "
    strData = strData & "1ECBBA07 14E2A97B 6143459D 5CDE5F14 53A8FBF1 D5F0AC89 97EA04D8 1C5E5B00 "
    strData = strData & "622181A8 D4FDB3F3 E9BCD341 572A1208 14112586 43CCE58A 9144C5FE A4C6E0A4 "
    strData = strData & "0D33D065 65CF620F 54A48D48 9F219CA1 C43E5EAC 6D63C821 A9728B3A 72770DAF "
    strData = strData & "D7934E7B 20DF87EF E35503B6 1A3E86E5 CAE321FB C819D504 129A50B3 AC60BFA6 "
    strData = strData & "CD5E68EA 7E9FB6C3 B01C9019 9483B1C7 3DE93CD5 C295376C AED52EDF 2AB9AD13 "
    strData = strData & "2E60F512 C0A07884 BC3D86A3 E36210C9 35269D9B 163951CE 0C7D6E2A D0CDB5FA "
    strData = strData & "59E86297 D87F5733 298EF221 898DB0E7 55000029 D1A5AA7E 8BC08AE1 B5061B45 "
    strData = strData & "C2C31C2B 6C92703A 94CC596B AF25EF42 0A1D73DB 22540456 04B6A0F9 D9C4179A "
    strData = strData & "EFFDAFA2 AE3D3C60 F7C8075B B49496C4 9CC5C714 1D1CD4E3 78BD1638 218E5534 "
    strData = strData & "B2F11568 F850246A EDFABCFA 9502BC29 796CE5F2 DA23051B AAE128B0 DC93537C "
    strData = strData & "3A493DA0 EE4B29AE B5DF6B2C 416895D7 FCABBD25 122D7F37 70810B58 105DC4B1 "
    strData = strData & "E10FDD37 F7882A90 524DCAB5 518A3F5C 3C9E8587 8451255B 40298281 19BD34E2 "
    strData = strData & "74A05B6F 5D3CECCB B6100215 42E13ECA 0FF979D1 2F59E2AC 6037DA27 E4F9CC50 "
    strData = strData & "5E92975A 0DF1847D D66DE190 D3E623FE 5032D6B8 7B568048 9A36B7CE 8235216E "
    strData = strData & "80272A7A 24F64B4A 93EFED8B 8C6916F7 37DDBFF4 4CCE1555 4B95DB5D 4B99BD25 "
    strData = strData & "92D3FDA1 69812FC0 FB1A4A9A 90660BB6 730C1969 46A4B9B2 81E289AA 7F49DA68 "
    strData = strData & "64669A0F 83B1A05F 27B3FF7D 9644F48B CC6B615C 8DB675B3 674F20B9 BCEBBE95 "
    strData = strData & "6F312382 75655982 5AE48871 3E45CF05 BF619F99 54C21157 EABAC460 40A8EAE9 "
    strData = strData & "454C6FE9 F2C0C1CD 419CF649 6412691C D3DC3BEF 265B0F70 6D0E60F5 C3578A9E "

    avntData = Split(strData, Chr$(32))   ' Load variant array

    ' load T3 table array
    For lngIndex = 0 To 511 Step 2

        ' Convert variant data to Long Integer
        T3(0, lngIdx) = CLng("&H" & avntData(lngIndex))
        T3(1, lngIdx) = CLng("&H" & avntData(lngIndex + 1))
        lngIdx = lngIdx + 1

    Next lngIndex

    strData = vbNullString   ' Verify string is empty
    avntData = Empty         ' Verify variant is empty
    lngIdx = 0

    ' Load T4 table array
    strData = strData & "5B0E6085 26323C55 1A46C1A9 FA1B59F5 A9E245A1 7C4C8FFA 65CA5159 DB2955D7 "
    strData = strData & "05DB0A76 CE35AFC2 81EAC77E A9113D45 528EF88A B6AC0A0D A09EA253 597BE3FF "
    strData = strData & "430DDFB3 AC48CD56 C4B3A67A F45CE46F 4ECECFD8 FBE2D05E 3EF56F10 B39935F0 "
    strData = strData & "0B22D682 9CD619C6 17FD460A 74DF2069 6CF8CC8E 8510ED40 D6C824BF 3A6ECAA7 "
    strData = strData & "61243D58 1A817049 048BACB6 BBC163A2 D9A38AC2 7D44CC32 7FDDFF5B AAF410AB "
    strData = strData & "AD6D495A A804824B E1A6A74F 2D8C9F94 D4F78512 35DEE8E3 FD4B7F88 6540D893 "
    strData = strData & "247C2004 2AA4BFDA 096EA1C5 17D1327C D56966B4 361A6685 277DA5C3 1221057D "
    strData = strData & "94D59893 A43ACFF7 64F0C51C CDC02281 3D33BCC4 FF6189DB E005CB18 4CE66AF1 "
    strData = strData & "FF5CCD1D 1DB99BEA B0B854A7 FE42980F 7BD46A6A 718D4B9F D10FA8CC 22A5FD8C "
    strData = strData & "D3148495 2BE4BD31 C7FA975F CB243847 4886ED1E 5846C407 28CDDB79 1EB70B04 "
    strData = strData & "C2B00BE2 F573417F 5C959045 2180F877 7A6BDDFF F370EB00 CE509E38 D6D9D6A4 "
    strData = strData & "EBEB0F00 647FA702 1DCC06CF 76606F06 E4D9F28B A286FF0A D85A305D C918C262 "
    strData = strData & "475B1D87 32225F54 2D4FB516 68CCB5FE A679B9D9 D72BBA20 53841C0D 912D43A5 "
    strData = strData & "3B7EAA48 BF12A4E8 781E0E47 F22F1DDF EFF20CE6 0AB50973 20D261D1 9DFFB742 "
    strData = strData & "16A12B03 062A2E39 1960EB22 39650495 251C16FE D50EB8B8 9AC0C330 F826016E "
    strData = strData & "ED152665 953E7671 02D63194 A6369570 5074F083 94B1C987 70BA598C 90B25CE1 "
    strData = strData & "794A1581 0B9742F6 0D5925E9 FCAF8C6C 3067716C D868744E 910AB077 E8D7731B "
    strData = strData & "6A61BBDB 5AC42F61 93513EFB F0851567 F494724B 9E83E9D5 E887E198 5C09648D "
    strData = strData & "34B1D3C6 75370CFD DC35E433 BC0D255D D0AAB842 34131BE0 08042A50 B48B7EAF "
    strData = strData & "9997C4EE 44A3AB35 829A7B49 201799D0 263B8307 B7C54441 752F95F4 FD6A6CA6 "
    strData = strData & "92721740 2C08C6E5 2A8AB754 A795D9EE A442F755 2F72943D 2C31334E 19781208 "
    strData = strData & "4FA98D7C EAEE6291 55C3862F 665DB309 BD061017 5D53B1F3 46FE6CB8 40413F27 "
    strData = strData & "3FE03792 DF0CFA59 CFE70037 2EB85E8F A7BE29E7 ADBCE118 E544EE5C DE8431DD "
    strData = strData & "8A781B1B 41F1873E A5C94C78 A0D2F0E7 39412E28 77B60728 A1265EF3 AFC9A62C "
    strData = strData & "BCC2770C 6A2506C5 3AB66DD5 DCE1CE12 E65499D0 4A675B37 7D8F5234 81BFD216 "
    strData = strData & "0F6F64FC EC15F389 74EFBE61 8B5B13C8 ACDC82B7 14273E1D DD40BFE0 03199D17 "
    strData = strData & "37E99257 E7E061F8 FA526269 04775AAA 8BBBF63A 463D56F9 F0013F15 43A26E64 "
    strData = strData & "A8307E9F 879EC898 CC4C27A4 150177CC 1B432F2C CA1D3348 DE1D1F8F 9F6FA013 "
    strData = strData & "606602A0 47A7DDD6 D237AB64 CC1CB2C7 9B938E72 25FCD1D3 EC4E0370 8E0FF476 "
    strData = strData & "FEB2FBDA 3D03C12D AE0BCED2 EE43889A 22CB8923 EBFB4F43 69360D01 3CF7396D "
    strData = strData & "855E3602 D2D4E022 073805BA D01F784C 33E17A13 3852F546 DF487405 8AC7B638 "
    strData = strData & "BA92B29C 678AA14A 0CE89FC7 6CFAADCD 5F9D4E09 08339E34 F1AFE929 1F5923B9 "
    strData = strData & "6E3480F6 0F4A265F EEBF3A2A B29B841C E21938A8 8F91B4AD 57DFEFF8 45C6D3C3 "
    strData = strData & "2F006B0B F62CAAF2 62F479EF 6F75EE78 11A55AD4 1C8916A9 F229D290 84FED453 "
    strData = strData & "42F1C27B 16B000E6 2B1F7674 9823C074 4B76ECA3 C2745360 8C98F463 B91691BD "
    strData = strData & "14BCC93C F1ADE66A 8885213E 6D458397 8E177DF0 274D4711 B49B73B5 503F2951 "
    strData = strData & "10168168 C3F96B6B 0E3D963B 63CAB0AE 8DFC4B56 55A1DB14 F789F135 6E14DE5C "
    strData = strData & "683E68AF 4E51DAC1 C9A84F9D 8D4B0FD9 3691E03F 52A0F9D1 5ED86E46 E1878E80 "
    strData = strData & "3C711A0E 99D07150 5A0865B2 0C4E9310 56FBFC1F E4F0682E EA8D5DE3 105EDF9B "
    strData = strData & "71ABFDB1 2379187A 2EB99DE1 BEE77B9C 21ECC0EA 33CF4523 59A4D752 1805C7A1 "
    strData = strData & "3896F5EB 56AE7C72 AA638F3D B18F75DC 9F39358D ABE9808E B7DEFA91 C00B72AC "
    strData = strData & "6B5541FD 62492D92 6DC6DEE8 F92E4D5B 353F57AB C4BEEA7E 735769D6 DA5690CE "
    strData = strData & "0A234AA6 42391484 F6F95080 28F80D9D B8E319A2 7AB3F215 31AD9C11 51341A4D "
    strData = strData & "773C22A5 7BEF5805 45C7561A 07968633 F913DA9E 249DBE36 DA652D9B 78A64C68 "
    strData = strData & "4C27A97F 3BC334EF 76621220 E66B17F4 96774389 9ACD7D0B F3EE5BCA E0ED6782 "
    strData = strData & "409F7536 00C879FC 06D09A39 B5926DB6 6F83AEB0 317AC588 01E6CA4A 86381F21 "
    strData = strData & "66FF3462 D19F3025 72207C24 DDFD3BFB 4AF6B6D3 E2ECE2EB 9C994DBE C7EA08DE "
    strData = strData & "49ACE597 B09A8BC4 B38C4766 CF0797BA 131B9373 C57C2A75 B1822CCE 61931E58 "
    strData = strData & "9D7555B9 09BA1C0C 127FAFDD 937D11D2 29DA3BAD C66D92E4 A2C1D571 54C2ECBC "
    strData = strData & "58C5134D 82F6FE24 1C3AE351 5B62274F E907C82E 01CB8126 F8ED0919 13E37FCB "
    strData = strData & "3249D8F9 C80046C9 80CF9BED E388FB63 1881539A 116CF19E 5103F3F7 6BD52457 "
    strData = strData & "15B7E6F5 AE47F7A8 DBD7C6DE D47E9CCF 44E55C41 0228BB1A B647D425 5EDB4E99 "
    strData = strData & "5D11882B B8AAFC30 F5098BBB 29D3212A 8FB5EA14 E90296B3 677B9421 57DD025A "
    strData = strData & "FB58E7C0 A390ACB5 89D3674C 83BD4A01 9E2DA4DF 4BF3B93B FCC41E32 8CAB4829 "
    strData = strData & "03F38C96 BA582C52 CAD1BDBD 7FD85DB2 BBB442C1 6082AE83 B95FE86B A5DA9AB0 "
    strData = strData & "B22E0467 3771A93F 845358C9 493152D8 BE2A4886 97B4541E 95A2DC2D D38E6966 "
    strData = strData & "C02C11AC 923C852B 2388B199 0DF2A87B 7C8008FA 1B4F37BE 1F70D0C8 4D54E503 "
    strData = strData & "5490ADEC 7ECE57D4 002B3C27 D9063A3A 7EAEA384 8030A2BF C602326D ED2003C0 "
    strData = strData & "83A7287D 69A94086 C57A5FCB 30F57A8A B56844E4 79EBE779 A373B40F 05DCBCE9 "
    strData = strData & "D71A786E 88570EE2 879CBACD BDE8F6A0 976AD1BC C164A32F AB21E25E 9666D78B "
    strData = strData & "901063AA E5E5C33C 9818B344 48698D90 E36487AE 3E1E8ABB AFBDF931 893BDCB4 "
    strData = strData & "6345A0DC 5FBBD519 8628FE26 9B9465CA 1E5D0160 3F9C51EC 4DE44006 A15049B7 "
    strData = strData & "BF6C70E5 F776CBB1 411218F2 EF552BED CB0C0708 705A36A3 E74D1475 4F986044 "
    strData = strData & "CD56D943 0EA8280E C12591D7 535F5065 C83223F1 720AEF96 C3A0396F 7363A51F "

    avntData = Split(strData, Chr$(32))   ' Load variant array

    ' load T4 table array
    For lngIndex = 0 To 511 Step 2

        ' Convert variant data to Long Integer
        T4(0, lngIdx) = CLng("&H" & avntData(lngIndex))
        T4(1, lngIdx) = CLng("&H" & avntData(lngIndex + 1))
        lngIdx = lngIdx + 1

    Next lngIndex

    strData = vbNullString   ' Verify string is empty
    avntData = Empty         ' Verify variant is empty

End Sub

Private Sub ResetVariables()

    ' Called by Class_Initialize()
    '           Class_Terminate()
    
    ' Empty arrays
    Erase T1()
    Erase T2()
    Erase T3()
    Erase T4()
    Erase malngHash()

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
    ' Enum enumTIGER3_ALGORITHM
    Const eTiger3_128 As Long = 0&
    Const eTiger3_160 As Long = 1&
    Const eTiger3_192 As Long = 2&
    Const eTiger3_224 As Long = 3&
    Const eTiger3_256 As Long = 4&
    Const eTiger3_384 As Long = 5&
    Const eTiger3_512 As Long = 6&
End Sub
#End If

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    ResetVariables
    Set mobjMath32 = Nothing      ' Free class objects from memory
    Set mobjMath64 = Nothing

    Set mobjMath32 = New cMath32  ' Instantiate class objects
    Set mobjMath64 = New cMath64

    ' Preset property values
    ReturnLowercase = False       ' Default return of hashed data
    HashRounds = 3                ' Set default number of passes
    HashMethod = eTiger3_192        ' Set default hash algorithm

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.

    ResetVariables
    Set mobjMath32 = Nothing   ' Free class objects from memory
    Set mobjMath64 = Nothing

End Sub

