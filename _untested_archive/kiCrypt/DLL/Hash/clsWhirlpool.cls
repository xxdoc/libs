VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWhirlpool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
'  Module:     cWhirlpool  (clsWhirlpool.cls)
'
'              Whirlpool versions for -224, -256, -384 bit output are my
'              experiment.  Any problems with these outputs, please email
'              me at:
'
'                        Kenneth Ives  kenaso@tx.rr.com
'
'              DO NOT CONTACT   Vincent Rijmen
'                               Paulo S.L.M.Barreto
'
'              because they did not write this module nor are they
'              responsible in any manner as to its content.
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
' Description:   The Whirlpool Hashing Function
'                Designed by Vincent Rijmen and Paulo S. L. M. Barreto
'
'                The Whirlpool hashing function is named after the Whirlpool
'                galaxy in Canes Venatici (M51, or NGC 5194), the first one
'                recognized to have spiral structure by William Parsons,
'                third Earl of Rosse, in April 1845 (cf. M. Hoskin, "The
'                Cambridge Illustrated History of Astronomy," Cambridge
'                University Press, 1997).
'
'                Whirlpool is a cryptographic hash function adopted by
'                the International Standards Organization (ISO) and
'                International Electrotechnical Commission (IEC) as part
'                of the joint ISO/IEC 10118-3 international standard.
'                It takes an arbritrary block of data and returns a 512
'                bit digest that can be used as a digital fingerprint for
'                message authentication.  Compliance with the standard
'                may be verified at:
'                http://hash.online-convert.com/whirlpool-generator
'
'                Whirlpool is a hash designed after the Square block cipher.
'                Whirlpool is a Miyaguchi-Preneel construction based on a
'                substantially modified Advanced Encryption Standard (AES).
'                It takes a message of any length less than 2256 bits and
'                returns a 512-bit message digest.
'
'                The authors have declared that "Whirlpool is not (and will
'                never be) patented and may be used free of charge for any
'                purpose. The reference implementations are in the public
'                domain."
'
'                Using the reference C implementation on a 1 GHz Pentium
'                III platform, we observe that Whirlpool operates at
'                about 73 cycles per hashed byte.
'
'                The compression function runs at about 56 cycles per
'                hashed byte. Many factors explain the observed
'                performance. First, a 32-bit processor was used to test
'                a native 64-bit implementation; better results are
'                expected by merely running the speed measurement on an
'                Alpha or Itanium processor. Second, it seems that the
'                pipe parallelism capabilities of the Pentium were not
'                fully used; this may reflect a non-optimising
'                implementation of 64-bit arithmetic support by the C
'                compiler, and might be overcome by an assembler
'                implementation. Third, the tables employed in the
'                reference implementation are quite large, and the
'                built-in processor cache might not be enough to hold
'                them, the data being hashed, and the hashing code at
'                once, thus degrading processing speed.
'
'                Whirlpool is much more scalable than most modern
'                hashing functions. Even though is not specifically
'                oriented toward any platform, it is rather efficient on
'                many of them, its structure favouring extensively
'                parallel execution of the component mappings. At the
'                same time, it does not require excessive storage space
'                (either for code or for tables), and can therefore be
'                efficiently implemented in quite constrained
'                environments like smart cards, although it can benefit
'                from larger cache memory available on modern processors
'                to achieve higher performance. It does not use
'                expensive or unusual instructions that must be built in
'                the processor. The mathematical simplicity of the
'                primitive resulting from the design strategy tends to
'                make analysis easier. And finally, it has a very long
'                hash length; this not only provides increased
'                protection against birthday attacks, but also offers a
'                larger internal state for entropy containment, as is
'                needed for certain classes of pseudo-random number
'                generators.
'
' Reference:     The Whirlpool Hash Function (Home page)
'                http://www.larc.usp.br/~pbarreto/WhirlpoolPage.html
'
'                Whirlpool (cryptography)
'                http://en.wikipedia.org/wiki/Whirlpool_(cryptography)
'
'                Whirlpool Hashing Function in Visual Basic 6.0
'                John Korejwa  <korejwa@tiac.net>
'                31-Dec-2010
'                http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=73638&lngWId=1
'
' ***************************************************************************
'
'  Whirlpool hash output using one (1) pass
'
'  Whirlpool-512 in accordance with ISO test vectors tested to be correct.
'  Whirlpool-224, -256, -384 are dynamic (unofficial) and can be changed
'  by the hash distributor.
'
'  "a"
'        8aca2602792aec6f11a67206531fb7d7f0dff59413145e6973c45001d0087b42d11bc645413aeff63a42391a39145a591a92200d560195e53b478584fdae231a
'
'  "abc"
'        4e2448a4c6f486bb16b6562c73b4020bf3043e3a731bce721ae1b303d97e6d4c7181eebdb6c57e277d0e34957114cbd6c797fc9d95d8b582d225292076d4eef5
'
'  "message digest"
'        378c84a4126e2dc6e56dcc7458377aac838d00032230f53ce1f5700c0ffb4d3b8421557659ef55c106b4b52ac5a4aaa692ed920052838f3362e86dbd37a8903e
'
'  "abcdefghijklmnopqrstuvwxyz"
'        f1d754662636ffe92c82ebb9212a484a8d38631ead4238f5442ee13b8054e41b08bf2a9251c30b6a0b8aae86177ab4a6f68f673e7207865d5d9819a3dba4eb3b
'
'  "abcdbcdecdefdefgefghfghighijhijk"
'        2a987ea40f917061f5d6f0a0e4644f488a7a5a52deee656207c562f988e95c6916bdc8031bc5be1b7b947639fe050b56939baaa0adff9ae6745b7b181c3be3fd
'
'  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'        dc37e008cf9ee69bf11f00ed9aba26901dd7c28cdec066cc6af42e40f82f3a1e08eba26629129d8fb7cb57211b9281a65517cc879d7b962142c65f5a7af01467
'
'  8 times "1234567890"
'        466ef18babb0154d25b9d38a6414f5c08784372bccb204d6549c4afadb6014294d5bd8df2a6c44e538cd047b2681a51a2c60481e88c5a20b2c2a80cf3a9a083b
'
'  1 million letter "a"
'        0c99005beb57eff50a7cf005560ddf5d29057fd86b20bfd62deca0f1ccea4af51fc15490eddc47af32bb2b66c34ff9ad8c6008ad677f77126953b226e4ed8b01
'
'  1 million binary zeroes
'        f3008d58b928890f6c902cc7e69f292cfd4870a015e7d4200612a5a3a7eb5b71be0bf894d2a030b59d7a608e63daf69cfd3d69cd77e446ebf4952b9470d820af
'
'  10 million binary zeroes  (Non ISO - exceeds 5mb limit in this module)
'        6d578ff856536058c50b8c737a71217126a0777e9a5c0443d10dfacd28da6c7f8988dfe02e325aa94b3ada8ae8b52a1dd199c3e47ce05efa673cc4182f63af1b
'
' ***************************************************************************
'
'  To create a 10mb test file, use the following code:
'
'     Dim hFile As Long
'     hFile = Freefile
'     Open "C:\Temp\MB_10.tst" For Binary Access Write As #hFile
'     Put #hFile, 10000000, Chr$(0)    ' Fill with binary 0's
'     Close #hFile
'
'  Got same results hashing 10mb file with this module and at web site
'  http://hash.online-convert.com/whirlpool-generator
'
'  Creating a file this way is fast and easy.  The CHR$(0) is the last
'  character in the file.  All previous characters are null values.  You
'  can create files up to one gigabyte in size.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 31-Dec-2010  John Korejwa  korejwa@tiac.net
'              Whirlpool Hashing Function in Visual Basic 6.0
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=73638&lngWId=1
' 05-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 30-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified to produce 256 and 384 bit hashes.  See LoadWorkArrays()
'              and FormatOutput() routines.
' 13-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified to produce 224 bit hashes.  See LoadWorkArrays() and
'              FormatOutput() routines.
' 20-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated number of data mixing iterations. See LoadWorkArrays()
'              routine.
' 04-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added new property ReturnLowercse().  Designates if returned
'              hashed data string should be in upper or lowercase format.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsWhirlpool"
  Private Const MAX_ROUNDS  As Long = 10
  Private Const BLOCK_SIZE  As Long = 64
  Private Const WORK_SIZE   As Long = 15
  Private Const BIT_CHUNK   As Long = 512
  Private Const MAX_BYTE    As Long = 256
  Private Const MB_5        As Long = &H500000   ' 5242880 bytes

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Private Enum enumWhirlpool_ALGORITHM
      eWhirlpool_224   ' 0
      eWhirlpool_256   ' 1
      eWhirlpool_384   ' 2
      eWhirlpool_512   ' 3  Default
  End Enum

' ***************************************************************************
' Module Variables
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Byte)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a byt Hashed
' Variable name:     mabytHashed
'
' ***************************************************************************
  Private mlngHashMethod    As enumWhirlpool_ALGORITHM
  Private mblnRetLowercase  As Boolean
  Private mcurAccrued       As Currency   ' Calculating progress
  Private mcurMaxSize       As Currency
  Private mlngHashRounds    As Long       ' Extend hash iterations (see property HashRounds)
  Private mlngTotalBitsLow  As Long       ' Accumulators
  Private mlngTotalBitsHigh As Long
  Private malngKey()        As Long       ' Round key
  Private malngHash()       As Long       ' Hashed data as long integer
  Private malngWord()       As Long       ' Data converted into long integers
  Private malngRound()      As Long       ' Round constants
  Private malngCTab0()      As Long       ' Circulant table constants
  Private malngCTab1()      As Long
  Private malngCTab2()      As Long
  Private malngCTab3()      As Long
  Private malngCTab4()      As Long
  Private malngCTab5()      As Long
  Private malngCTab6()      As Long
  Private malngCTab7()      As Long
  Private mabytHashed()     As Byte       ' Hashed results

' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let HashMethod(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eWhirlpool_512
    Select Case lngData
           Case 0 To 3: mlngHashMethod = lngData
           Case Else:   mlngHashMethod = eWhirlpool_512  ' Use default value
    End Select

End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngHashRounds = lngData   ' Good selection
           Case Else:            mlngHashRounds = 1         ' Default value
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
' Description:   Function to create a unique hex string representation of 
'                the data passed.  Can process files larger than 2 gb.
'                See Readme.txt file for more information. 
'
' Syntax:
'    With mobjWhirlpool
'        abytData() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'        mabytHashed() = .HashFile(abytData())                        ' hash data and return as Byte array
'        txtData.Text = StrConv(mabytHashed(), vbUnicode)             ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - Path/filename in byte array to be hashed
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function HashFile(ByRef abytFile() As Byte) As Byte()

    Dim hFile        As Long
    Dim lngByteCnt   As Long
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

    mcurAccrued = 0@  ' Init accrued progress
    curFilePos = 0@   ' set to first position in the file

    Set objBigFiles = New cBigFiles  ' Instantiate class object
    ResetVariables                   ' Reset class variables
    LoadWorkArrays                   ' load work and constant arrays

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    With objBigFiles

        strSource = ByteArrayToString(abytFile())   ' Convert byte array to string
        .CalcFileSize strSource, curAmtLeft         ' Get the size of the file
        mcurMaxSize = curAmtLeft                    ' Capture file size

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

        ' If file exceeds 5mb then
        ' process data in chunks
        Do While (curAmtLeft > MB_5)

            lngBlockSize = MB_5                        ' Adjust block size
            ReDim abytData(lngBlockSize - 1)               ' Size data array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)   ' Adjust amount of file left to process

            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)
            lngByteCnt = 0

            Do While (lngByteCnt < lngBlockSize)

                CopyMemory malngWord(0), abytData(lngByteCnt), BLOCK_SIZE   ' Load word array
                lngByteCnt = lngByteCnt + BLOCK_SIZE                        ' Increment index pointer
                mlngTotalBitsLow = mlngTotalBitsLow + BIT_CHUNK             ' Increment bit count
                Transform                                                   ' Hash data

                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do   ' exit Do..Loop
                End If

            Loop

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

            ' Update accumulators
            mlngTotalBitsHigh = mlngTotalBitsHigh + (mlngTotalBitsLow \ &H1000000)
            mlngTotalBitsLow = mlngTotalBitsLow And &HFFFFFF

        Loop

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo HashFile_CleanUp
        End If

        ' See if any data is left
        ' or file is less than 5mb
        If (curAmtLeft > 0@) Then

            lngBlockSize = CLng(curAmtLeft)    ' Final block size
            Erase abytData()                   ' empty data array
            ReDim abytData(lngBlockSize - 1)   ' Size receiving array

            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                GoTo HashFile_CleanUp
            End If

            ' Process message data
            If Not PrepareMsgData(abytData()) Then
                gblnStopProcessing = True
            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo HashFile_CleanUp
            End If

        End If

        .API_CloseFile hFile   ' Close file opened by this routine

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
    objBigFiles.API_CloseFile hFile  ' Close file opened by this routine
    Set objBigFiles = Nothing        ' Always free objects from memory
    ResetVariables                   ' Reset class variables

End Function

' ***************************************************************************
' Routine:       HashString
'
' Description:   Function to digest a text string and output the result as
'                a string of hexadecimal characters.  Got basic idea from
'                David Ireland of DI Management Services Pty Ltd
'                <www.di-mgt.com.au>.
'
' Syntax:
'    With mobjWhirlpool
'        abytData() = StrConv("abc", vbFromUnicode)       ' convert string data to byte array 
'        mabytHashed() = .HashString(abytData())          ' hash data and return as Byte array
'        txtData.Text = StrConv(mabytHashed(), vbUnicode) ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - Path/filename in byte array to be hashed
'
' Returns:       Hashed data string in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
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

    ResetVariables   ' Reset class variables
    LoadWorkArrays   ' load work and constant arrays

    mcurAccrued = 0@                                             ' Init accrued progress
    mcurMaxSize = CCur((UBound(abytData) + 1) * mlngHashRounds)  ' Capture data size

    ' Process message data
    If Not PrepareMsgData(abytData()) Then
        gblnStopProcessing = True    ' Set flag to stop processing
    Else
        HashString = FormatOutput()  ' Format final output
    End If

HashString_CleanUp:
    ResetVariables   ' Reset class variables

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

' ***************************************************************************
' Routine:       FormatOutput
'
' Description:   Concatenate the hashed values into one string.
'
'                Output length in hex format:
'                    Whirlpool-224 =  56 characters
'                    Whirlpool-256 =  64 characters
'                    Whirlpool-384 =  96 characters
'                    Whirlpool-512 = 128 characters
'
' Returns:       Hex data in byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 30-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified to produce 256 and 384 bit hashes.
' 13-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified to produce 224 bit hash.
' ***************************************************************************
Private Function FormatOutput() As Byte()

    Dim strOutput  As String
    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long

    strOutput = Space$(256)  ' preload output buffer
    lngPointer = 1           ' Start position in output buffer

    For lngIndex = 0 To BLOCK_SIZE - 1
        Mid$(strOutput, lngPointer, 2) = Right$("0" & Hex$(mabytHashed(lngIndex)), 2)
        lngPointer = lngPointer + 2
    Next lngIndex

    ' Adjust according to desired length in bytes (characters)
    Select Case mlngHashMethod
           Case eWhirlpool_224: lngLength = 56
           Case eWhirlpool_256: lngLength = 64
           Case eWhirlpool_384: lngLength = 96
           Case eWhirlpool_512: lngLength = 128
    End Select

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

Private Function PrepareMsgData(ByRef abytData() As Byte) As Boolean

    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim lngBitCnt   As Long
    Dim lngBitValue As Long
    Dim abytTemp()  As Byte

    Const ROUTINE_NAME As String = "PrepareMsgData"

    On Error GoTo PrepareMsgData_Error

    PrepareMsgData = False   ' Preset to FALSE

    ReDim abytTemp(BLOCK_SIZE - 1)  ' Size temp work array

    lngBitCnt = (UBound(abytData) + 1) * 8   ' Get bit count for incoming data

    mlngTotalBitsLow = mlngTotalBitsLow + lngBitCnt                         ' Increment accumulator
    mlngTotalBitsHigh = mlngTotalBitsHigh + (mlngTotalBitsLow \ &H1000000)  ' Increment accumulator
    mlngTotalBitsLow = mlngTotalBitsLow And &HFFFFFF                        ' Adjust low side accumulator
    lngIndex = 0                                                            ' Set array pointer

    Do While lngBitCnt >= BIT_CHUNK

        CopyMemory malngWord(0), abytData(lngIndex), BLOCK_SIZE   ' Load word array
        lngIndex = lngIndex + BLOCK_SIZE                          ' Increment index pointer
        lngBitCnt = lngBitCnt - BIT_CHUNK                         ' Decrement bit count
        Transform                                                 ' Hash data

    Loop

    lngIdx = lngBitCnt \ 8   ' Add remaining 8 bit chunks

    If lngIdx > 0 Then
        CopyMemory abytTemp(0), abytData(lngIndex), lngIdx
    End If

    lngBitCnt = lngBitCnt - 8 * lngIdx  ' Adjust bit counter
    lngBitValue = &H80&                 ' Add remaining bits
    abytTemp(lngIdx) = 0

    Do While lngBitCnt > 0

        If (abytData(lngIndex) And lngBitValue) Then
            abytTemp(lngIdx) = abytTemp(lngIdx) Or lngBitValue
        End If

        lngBitValue = lngBitValue \ 2
        lngBitCnt = lngBitCnt - 1

    Loop

    abytTemp(lngIdx) = abytTemp(lngIdx) Or lngBitValue    ' Append "1" bit
    lngIdx = lngIdx + 1                                   ' Increment pointer

    If lngIdx >= (BLOCK_SIZE \ 2) Then

        ' clear the rest of abytTemp()
        Do While lngIdx < (BLOCK_SIZE - 1)
            abytTemp(lngIdx) = 0
            lngIdx = lngIdx + 1
        Loop

        CopyMemory malngWord(0), abytTemp(0), BLOCK_SIZE   ' Load word array
        Transform                                          ' Hash data

        ' Reset temp array values
        For lngIdx = 0 To BLOCK_SIZE - 1
            abytTemp(lngIdx) = 0
        Next lngIdx

    Else

        ' Clear rest of temp array
        Do While lngIdx < BLOCK_SIZE
            abytTemp(lngIdx) = 0
            lngIdx = lngIdx + 1
        Loop

    End If

    ' append (right-justified) bit length of entire message
    abytTemp(63) = mlngTotalBitsLow And &HFF&
    abytTemp(62) = (mlngTotalBitsLow \ &H100&) And &HFF&
    abytTemp(61) = (mlngTotalBitsLow \ &H10000) And &HFF&
    abytTemp(60) = mlngTotalBitsHigh And &HFF&
    abytTemp(59) = (mlngTotalBitsHigh \ &H100&) And &HFF&
    abytTemp(58) = (mlngTotalBitsHigh \ &H10000) And &HFF&

    CopyMemory malngWord(0), abytTemp(0), BLOCK_SIZE  ' Load word array
    Transform                                         ' Hash data

    ' Store hashed results in byte array
    CopyMemory mabytHashed(0), malngHash(0), BLOCK_SIZE

    ' Reset hash data array
    For lngIndex = 0 To WORK_SIZE
        malngHash(lngIndex) = 0
    Next lngIndex

    PrepareMsgData = True   ' Good finish

PrepareMsgData_CleanUp:
    On Error GoTo 0      ' Nullify error trap in this routine
    Exit Function

PrepareMsgData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume PrepareMsgData_CleanUp

End Function

Private Sub Transform()

    Dim lngIndex    As Long
    Dim lngRounds   As Long
    Dim lngProgress As Long
    Dim HH()        As Long   ' Cipher state
    Dim abytIdx()   As Byte   ' Index pointer

    Const ROUTINE_NAME As String = "Transform"

    On Error GoTo Transform_Error

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo Transform_CleanUp
    End If

    ReDim abytIdx(BLOCK_SIZE)    ' Size work arrays
    ReDim HH(WORK_SIZE)

    CopyMemory malngKey(0), malngHash(0), BLOCK_SIZE  ' Load key array
    mcurAccrued = mcurAccrued + BLOCK_SIZE            ' Increment accumulator

    ' Load state array
    For lngIndex = 0 To WORK_SIZE
        HH(lngIndex) = malngWord(lngIndex) Xor malngHash(lngIndex)
    Next lngIndex

    ' For added security, use multiple iterations
    ' by updating property HashRounds()
    ' Default = 1 iteration
    For lngRounds = 1 To mlngHashRounds

        For lngIndex = 0 To 9

            CopyMemory abytIdx(0), malngKey(0), BLOCK_SIZE  ' Load index pointer array

            malngKey(0) = malngCTab0(0, abytIdx(0)) Xor malngCTab1(0, abytIdx(57)) Xor _
                          malngCTab2(0, abytIdx(50)) Xor malngCTab3(0, abytIdx(43)) Xor _
                          malngCTab4(0, abytIdx(36)) Xor malngCTab5(0, abytIdx(29)) Xor _
                          malngCTab6(0, abytIdx(22)) Xor malngCTab7(0, abytIdx(15)) Xor malngRound(0, lngIndex)

            malngKey(1) = malngCTab0(1, abytIdx(0)) Xor malngCTab1(1, abytIdx(57)) Xor _
                          malngCTab2(1, abytIdx(50)) Xor malngCTab3(1, abytIdx(43)) Xor _
                          malngCTab4(1, abytIdx(36)) Xor malngCTab5(1, abytIdx(29)) Xor _
                          malngCTab6(1, abytIdx(22)) Xor malngCTab7(1, abytIdx(15)) Xor malngRound(1, lngIndex)

            malngKey(2) = malngCTab0(0, abytIdx(8)) Xor malngCTab1(0, abytIdx(1)) Xor _
                          malngCTab2(0, abytIdx(58)) Xor malngCTab3(0, abytIdx(51)) Xor _
                          malngCTab4(0, abytIdx(44)) Xor malngCTab5(0, abytIdx(37)) Xor _
                          malngCTab6(0, abytIdx(30)) Xor malngCTab7(0, abytIdx(23))

            malngKey(3) = malngCTab0(1, abytIdx(8)) Xor malngCTab1(1, abytIdx(1)) Xor _
                          malngCTab2(1, abytIdx(58)) Xor malngCTab3(1, abytIdx(51)) Xor _
                          malngCTab4(1, abytIdx(44)) Xor malngCTab5(1, abytIdx(37)) Xor _
                          malngCTab6(1, abytIdx(30)) Xor malngCTab7(1, abytIdx(23))

            malngKey(4) = malngCTab0(0, abytIdx(16)) Xor malngCTab1(0, abytIdx(9)) Xor _
                          malngCTab2(0, abytIdx(2)) Xor malngCTab3(0, abytIdx(59)) Xor _
                          malngCTab4(0, abytIdx(52)) Xor malngCTab5(0, abytIdx(45)) Xor _
                          malngCTab6(0, abytIdx(38)) Xor malngCTab7(0, abytIdx(31))

            malngKey(5) = malngCTab0(1, abytIdx(16)) Xor malngCTab1(1, abytIdx(9)) Xor _
                          malngCTab2(1, abytIdx(2)) Xor malngCTab3(1, abytIdx(59)) Xor _
                          malngCTab4(1, abytIdx(52)) Xor malngCTab5(1, abytIdx(45)) Xor _
                          malngCTab6(1, abytIdx(38)) Xor malngCTab7(1, abytIdx(31))

            malngKey(6) = malngCTab0(0, abytIdx(24)) Xor malngCTab1(0, abytIdx(17)) Xor _
                          malngCTab2(0, abytIdx(10)) Xor malngCTab3(0, abytIdx(3)) Xor _
                          malngCTab4(0, abytIdx(60)) Xor malngCTab5(0, abytIdx(53)) Xor _
                          malngCTab6(0, abytIdx(46)) Xor malngCTab7(0, abytIdx(39))

            malngKey(7) = malngCTab0(1, abytIdx(24)) Xor malngCTab1(1, abytIdx(17)) Xor _
                          malngCTab2(1, abytIdx(10)) Xor malngCTab3(1, abytIdx(3)) Xor _
                          malngCTab4(1, abytIdx(60)) Xor malngCTab5(1, abytIdx(53)) Xor _
                          malngCTab6(1, abytIdx(46)) Xor malngCTab7(1, abytIdx(39))

            malngKey(8) = malngCTab0(0, abytIdx(32)) Xor malngCTab1(0, abytIdx(25)) Xor _
                          malngCTab2(0, abytIdx(18)) Xor malngCTab3(0, abytIdx(11)) Xor _
                          malngCTab4(0, abytIdx(4)) Xor malngCTab5(0, abytIdx(61)) Xor _
                          malngCTab6(0, abytIdx(54)) Xor malngCTab7(0, abytIdx(47))

            malngKey(9) = malngCTab0(1, abytIdx(32)) Xor malngCTab1(1, abytIdx(25)) Xor _
                          malngCTab2(1, abytIdx(18)) Xor malngCTab3(1, abytIdx(11)) Xor _
                          malngCTab4(1, abytIdx(4)) Xor malngCTab5(1, abytIdx(61)) Xor _
                          malngCTab6(1, abytIdx(54)) Xor malngCTab7(1, abytIdx(47))

            malngKey(10) = malngCTab0(0, abytIdx(40)) Xor malngCTab1(0, abytIdx(33)) Xor _
                           malngCTab2(0, abytIdx(26)) Xor malngCTab3(0, abytIdx(19)) Xor _
                           malngCTab4(0, abytIdx(12)) Xor malngCTab5(0, abytIdx(5)) Xor _
                           malngCTab6(0, abytIdx(62)) Xor malngCTab7(0, abytIdx(55))

            malngKey(11) = malngCTab0(1, abytIdx(40)) Xor malngCTab1(1, abytIdx(33)) Xor _
                           malngCTab2(1, abytIdx(26)) Xor malngCTab3(1, abytIdx(19)) Xor _
                           malngCTab4(1, abytIdx(12)) Xor malngCTab5(1, abytIdx(5)) Xor _
                           malngCTab6(1, abytIdx(62)) Xor malngCTab7(1, abytIdx(55))

            malngKey(12) = malngCTab0(0, abytIdx(48)) Xor malngCTab1(0, abytIdx(41)) Xor _
                           malngCTab2(0, abytIdx(34)) Xor malngCTab3(0, abytIdx(27)) Xor _
                           malngCTab4(0, abytIdx(20)) Xor malngCTab5(0, abytIdx(13)) Xor _
                           malngCTab6(0, abytIdx(6)) Xor malngCTab7(0, abytIdx(63))

            malngKey(13) = malngCTab0(1, abytIdx(48)) Xor malngCTab1(1, abytIdx(41)) Xor _
                           malngCTab2(1, abytIdx(34)) Xor malngCTab3(1, abytIdx(27)) Xor _
                           malngCTab4(1, abytIdx(20)) Xor malngCTab5(1, abytIdx(13)) Xor _
                           malngCTab6(1, abytIdx(6)) Xor malngCTab7(1, abytIdx(63))

            malngKey(14) = malngCTab0(0, abytIdx(56)) Xor malngCTab1(0, abytIdx(49)) Xor _
                           malngCTab2(0, abytIdx(42)) Xor malngCTab3(0, abytIdx(35)) Xor _
                           malngCTab4(0, abytIdx(28)) Xor malngCTab5(0, abytIdx(21)) Xor _
                           malngCTab6(0, abytIdx(14)) Xor malngCTab7(0, abytIdx(7))

            malngKey(15) = malngCTab0(1, abytIdx(56)) Xor malngCTab1(1, abytIdx(49)) Xor _
                           malngCTab2(1, abytIdx(42)) Xor malngCTab3(1, abytIdx(35)) Xor _
                           malngCTab4(1, abytIdx(28)) Xor malngCTab5(1, abytIdx(21)) Xor _
                           malngCTab6(1, abytIdx(14)) Xor malngCTab7(1, abytIdx(7))

            CopyMemory abytIdx(0), HH(0), BLOCK_SIZE  ' Load index pointer array

            HH(0) = malngCTab0(0, abytIdx(0)) Xor malngCTab1(0, abytIdx(57)) Xor _
                    malngCTab2(0, abytIdx(50)) Xor malngCTab3(0, abytIdx(43)) Xor _
                    malngCTab4(0, abytIdx(36)) Xor malngCTab5(0, abytIdx(29)) Xor _
                    malngCTab6(0, abytIdx(22)) Xor malngCTab7(0, abytIdx(15)) Xor malngKey(0)

            HH(1) = malngCTab0(1, abytIdx(0)) Xor malngCTab1(1, abytIdx(57)) Xor _
                    malngCTab2(1, abytIdx(50)) Xor malngCTab3(1, abytIdx(43)) Xor _
                    malngCTab4(1, abytIdx(36)) Xor malngCTab5(1, abytIdx(29)) Xor _
                    malngCTab6(1, abytIdx(22)) Xor malngCTab7(1, abytIdx(15)) Xor malngKey(1)

            HH(2) = malngCTab0(0, abytIdx(8)) Xor malngCTab1(0, abytIdx(1)) Xor _
                    malngCTab2(0, abytIdx(58)) Xor malngCTab3(0, abytIdx(51)) Xor _
                    malngCTab4(0, abytIdx(44)) Xor malngCTab5(0, abytIdx(37)) Xor _
                    malngCTab6(0, abytIdx(30)) Xor malngCTab7(0, abytIdx(23)) Xor malngKey(2)

            HH(3) = malngCTab0(1, abytIdx(8)) Xor malngCTab1(1, abytIdx(1)) Xor _
                    malngCTab2(1, abytIdx(58)) Xor malngCTab3(1, abytIdx(51)) Xor _
                    malngCTab4(1, abytIdx(44)) Xor malngCTab5(1, abytIdx(37)) Xor _
                    malngCTab6(1, abytIdx(30)) Xor malngCTab7(1, abytIdx(23)) Xor malngKey(3)

            HH(4) = malngCTab0(0, abytIdx(16)) Xor malngCTab1(0, abytIdx(9)) Xor _
                    malngCTab2(0, abytIdx(2)) Xor malngCTab3(0, abytIdx(59)) Xor _
                    malngCTab4(0, abytIdx(52)) Xor malngCTab5(0, abytIdx(45)) Xor _
                    malngCTab6(0, abytIdx(38)) Xor malngCTab7(0, abytIdx(31)) Xor malngKey(4)

            HH(5) = malngCTab0(1, abytIdx(16)) Xor malngCTab1(1, abytIdx(9)) Xor _
                    malngCTab2(1, abytIdx(2)) Xor malngCTab3(1, abytIdx(59)) Xor _
                    malngCTab4(1, abytIdx(52)) Xor malngCTab5(1, abytIdx(45)) Xor _
                    malngCTab6(1, abytIdx(38)) Xor malngCTab7(1, abytIdx(31)) Xor malngKey(5)

            HH(6) = malngCTab0(0, abytIdx(24)) Xor malngCTab1(0, abytIdx(17)) Xor _
                    malngCTab2(0, abytIdx(10)) Xor malngCTab3(0, abytIdx(3)) Xor _
                    malngCTab4(0, abytIdx(60)) Xor malngCTab5(0, abytIdx(53)) Xor _
                    malngCTab6(0, abytIdx(46)) Xor malngCTab7(0, abytIdx(39)) Xor malngKey(6)

            HH(7) = malngCTab0(1, abytIdx(24)) Xor malngCTab1(1, abytIdx(17)) Xor _
                    malngCTab2(1, abytIdx(10)) Xor malngCTab3(1, abytIdx(3)) Xor _
                    malngCTab4(1, abytIdx(60)) Xor malngCTab5(1, abytIdx(53)) Xor _
                    malngCTab6(1, abytIdx(46)) Xor malngCTab7(1, abytIdx(39)) Xor malngKey(7)

            HH(8) = malngCTab0(0, abytIdx(32)) Xor malngCTab1(0, abytIdx(25)) Xor _
                    malngCTab2(0, abytIdx(18)) Xor malngCTab3(0, abytIdx(11)) Xor _
                    malngCTab4(0, abytIdx(4)) Xor malngCTab5(0, abytIdx(61)) Xor _
                    malngCTab6(0, abytIdx(54)) Xor malngCTab7(0, abytIdx(47)) Xor malngKey(8)

            HH(9) = malngCTab0(1, abytIdx(32)) Xor malngCTab1(1, abytIdx(25)) Xor _
                    malngCTab2(1, abytIdx(18)) Xor malngCTab3(1, abytIdx(11)) Xor _
                    malngCTab4(1, abytIdx(4)) Xor malngCTab5(1, abytIdx(61)) Xor _
                    malngCTab6(1, abytIdx(54)) Xor malngCTab7(1, abytIdx(47)) Xor malngKey(9)

            HH(10) = malngCTab0(0, abytIdx(40)) Xor malngCTab1(0, abytIdx(33)) Xor _
                     malngCTab2(0, abytIdx(26)) Xor malngCTab3(0, abytIdx(19)) Xor _
                     malngCTab4(0, abytIdx(12)) Xor malngCTab5(0, abytIdx(5)) Xor _
                     malngCTab6(0, abytIdx(62)) Xor malngCTab7(0, abytIdx(55)) Xor malngKey(10)

            HH(11) = malngCTab0(1, abytIdx(40)) Xor malngCTab1(1, abytIdx(33)) Xor _
                     malngCTab2(1, abytIdx(26)) Xor malngCTab3(1, abytIdx(19)) Xor _
                     malngCTab4(1, abytIdx(12)) Xor malngCTab5(1, abytIdx(5)) Xor _
                     malngCTab6(1, abytIdx(62)) Xor malngCTab7(1, abytIdx(55)) Xor malngKey(11)

            HH(12) = malngCTab0(0, abytIdx(48)) Xor malngCTab1(0, abytIdx(41)) Xor _
                     malngCTab2(0, abytIdx(34)) Xor malngCTab3(0, abytIdx(27)) Xor _
                     malngCTab4(0, abytIdx(20)) Xor malngCTab5(0, abytIdx(13)) Xor _
                     malngCTab6(0, abytIdx(6)) Xor malngCTab7(0, abytIdx(63)) Xor malngKey(12)

            HH(13) = malngCTab0(1, abytIdx(48)) Xor malngCTab1(1, abytIdx(41)) Xor _
                     malngCTab2(1, abytIdx(34)) Xor malngCTab3(1, abytIdx(27)) Xor _
                     malngCTab4(1, abytIdx(20)) Xor malngCTab5(1, abytIdx(13)) Xor _
                     malngCTab6(1, abytIdx(6)) Xor malngCTab7(1, abytIdx(63)) Xor malngKey(13)

            HH(14) = malngCTab0(0, abytIdx(56)) Xor malngCTab1(0, abytIdx(49)) Xor _
                     malngCTab2(0, abytIdx(42)) Xor malngCTab3(0, abytIdx(35)) Xor _
                     malngCTab4(0, abytIdx(28)) Xor malngCTab5(0, abytIdx(21)) Xor _
                     malngCTab6(0, abytIdx(14)) Xor malngCTab7(0, abytIdx(7)) Xor malngKey(14)

            HH(15) = malngCTab0(1, abytIdx(56)) Xor malngCTab1(1, abytIdx(49)) Xor _
                     malngCTab2(1, abytIdx(42)) Xor malngCTab3(1, abytIdx(35)) Xor _
                     malngCTab4(1, abytIdx(28)) Xor malngCTab5(1, abytIdx(21)) Xor _
                     malngCTab6(1, abytIdx(14)) Xor malngCTab7(1, abytIdx(7)) Xor malngKey(15)

        Next lngIndex

    Next lngRounds

    ' apply the Miyaguchi-Preneel compression function
    For lngIndex = 0 To WORK_SIZE
        malngHash(lngIndex) = malngHash(lngIndex) Xor HH(lngIndex) Xor malngWord(lngIndex)
        malngWord(lngIndex) = 0
    Next lngIndex

    ' Update progress bar
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent HashProgress(lngProgress)

Transform_CleanUp:
    On Error GoTo 0
    Exit Sub

Transform_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Transform_CleanUp

End Sub

Private Sub LoadWorkArrays()

    Dim strData    As String
    Dim avntData   As Variant
    Dim lngIdx     As Long
    Dim lngIndex   As Long
    Dim abytBase() As Byte
    Dim abytTemp() As Byte

    On Error GoTo LoadArrays_Error

    ' Size arrays
    ReDim mabytHashed(BLOCK_SIZE)  ' Hashed output
    ReDim malngKey(WORK_SIZE)      ' Round key
    ReDim malngHash(WORK_SIZE)     ' Hashed data as long integer
    ReDim malngWord(WORK_SIZE)     ' Data converted into long integers
    ReDim malngRound(0 To 1, 9)    ' Round constants

    ReDim malngCTab0(0 To 1, MAX_BYTE)  ' Circulant tables
    ReDim malngCTab1(0 To 1, MAX_BYTE)
    ReDim malngCTab2(0 To 1, MAX_BYTE)
    ReDim malngCTab3(0 To 1, MAX_BYTE)
    ReDim malngCTab4(0 To 1, MAX_BYTE)
    ReDim malngCTab5(0 To 1, MAX_BYTE)
    ReDim malngCTab6(0 To 1, MAX_BYTE)
    ReDim malngCTab7(0 To 1, MAX_BYTE)

    avntData = Empty  ' Always start with empty variants

    ReDim abytBase(MAX_BYTE)   ' Size work arrays
    ReDim abytTemp(WORK_SIZE)

    strData = vbNullString
    mlngTotalBitsLow = 0    ' Empty accumulators
    mlngTotalBitsHigh = 0

    ' 01-Jan-2012 Bug fix - Duplicated a character in original data while
    '             attempting to make it look cleaner
    '
    ' These are ASCII values 0-255 mixed in hex format with no duplicates.
    ' See Whirlpool_tbl.txt for additional sets of data strings from which
    ' to choose.  See application BuildTables.vbp to create new data.
    '
    ' Whirlpool-224, -256, -384 are dynamic and can be modified/changed
    ' by the hash distributor.
    '
    ' Each case statement:
    '    1.  Consists of all ASCII values (0-255) in two
    '        character hex equivalent with no duplicates
    '    2.  Mixed 100-400 iterations using Knuth Shuffle
    Select Case mlngHashMethod
           Case eWhirlpool_224
                strData = strData & "E3 9E F1 8D 0D 66 14 51 CC CB 6C 56 60 27 AB FD "
                strData = strData & "BE B9 91 FB EB 48 2C 01 DE 74 45 A6 D5 64 10 31 "
                strData = strData & "A4 3E EC 7B F4 0C FE 9B 4E 94 3C AD 92 97 77 4F "
                strData = strData & "11 5E 95 0A DD FA 71 83 1A D8 67 8A C4 8C EF F8 "
                strData = strData & "54 7D 22 E7 5D D1 E9 88 42 D9 65 F9 28 D4 E2 47 "
                strData = strData & "E5 2E EA 1E AE 9F 41 0E DC 5C 0F CE 98 A2 D2 21 "
                strData = strData & "62 D0 93 85 A7 68 E1 6A 18 A5 26 F2 1D FF 34 6D "
                strData = strData & "96 5B 32 E6 30 6F A8 08 D3 4B 99 BD C3 CA 40 A3 "
                strData = strData & "17 57 C7 3B 38 16 50 F0 07 BC 05 15 82 87 53 29 "
                strData = strData & "9A 1B 8F 2B D6 9D 58 89 86 CD 04 7A C5 DF 78 DB "
                strData = strData & "B1 EE 13 3F 20 6E AA 1C 75 72 2D E0 81 B8 B6 5A "
                strData = strData & "0B CF D7 7F 36 4A F7 B2 A1 02 55 80 19 F5 DA E4 "
                strData = strData & "44 C0 06 BB 79 09 3A BF 24 FC 7C C9 F6 52 C2 61 "
                strData = strData & "12 B4 25 73 00 63 1F B7 23 B5 E8 33 4D 76 5F 9C "
                strData = strData & "70 ED 4C F3 7E A9 46 43 C6 8E 35 A0 AC 37 03 B0 "
                strData = strData & "BA 90 59 49 84 AF 39 2F 69 3D C8 6B C1 2A B3 8B "

           Case eWhirlpool_256
                strData = strData & "FD 45 E5 96 17 F8 6F AA 3A CA 10 87 90 66 8E 4C "
                strData = strData & "58 77 A0 57 1A 93 BA 22 9A 5D F0 32 B8 8F 1F E0 "
                strData = strData & "E2 F1 7D 28 FF FA D9 A7 4D 35 AC DE 47 50 A5 F2 "
                strData = strData & "67 61 C2 4A 11 88 73 3B D4 04 82 E9 9C 71 3E A9 "
                strData = strData & "70 7E 89 69 38 8B C0 1D 62 27 3D 39 AB 75 1B 59 "
                strData = strData & "2B 64 EF D7 25 D0 13 7B 4F C3 63 3F D1 91 23 C6 "
                strData = strData & "8A 9D BD 2D 08 30 2A 31 0B E4 5A F9 C8 94 09 0D "
                strData = strData & "ED 03 97 5E 42 54 99 FC 72 07 C5 EC 49 26 BC DA "
                strData = strData & "FE CB 6C 81 C7 9F 44 29 D8 21 85 AF 24 EE 0E 74 "
                strData = strData & "60 9E 2E 8D 80 EA 5C 79 4E DB 06 DF 05 41 7A F7 "
                strData = strData & "8C 68 16 F5 48 EB AD 7C 37 19 33 C9 20 4B A1 DD "
                strData = strData & "15 6B 0C BF 02 12 C1 D6 7F B0 E6 E8 5F B6 9B FB "
                strData = strData & "43 E1 CD 6A B5 E7 F6 BE 40 B9 A2 46 CE 00 01 18 "
                strData = strData & "CF BB B2 B1 6E 2C 98 F3 5B 92 56 DC A6 65 2F D2 "
                strData = strData & "53 D5 E3 36 6D B4 B7 51 78 34 A4 86 83 C4 AE 3C "
                strData = strData & "0A 0F A3 55 B3 76 52 A8 CC 1C 14 F4 95 84 1E D3 "

           Case eWhirlpool_384
                strData = strData & "CC 52 DE 47 BC 63 4F 0F 5B 9A 95 6E AC 78 88 83 "
                strData = strData & "FB F2 1B E4 62 D9 B0 F6 21 11 CF 8D 3D 64 A8 C4 "
                strData = strData & "D4 D8 13 91 F7 59 29 57 07 AB 32 94 CD E8 33 2A "
                strData = strData & "9B 0E 1F 42 7F 93 B9 43 1C 28 9D B3 D1 F3 41 69 "
                strData = strData & "CB C5 46 3E 06 FE 4D 89 EC E0 E7 8E 27 50 BF 7D "
                strData = strData & "2C 82 10 D6 71 E6 70 0C C1 5E FF A3 26 BB 02 D3 "
                strData = strData & "68 AE 39 66 2D 4C 60 3B 1E 73 C6 A9 FA 65 B2 C3 "
                strData = strData & "38 05 B4 E3 7E AF D7 84 6D FC 8B C8 E9 81 7B AD "
                strData = strData & "BA 8A D2 61 2E 35 01 4A 2F DA C7 34 EE F9 15 0D "
                strData = strData & "86 A6 AA DC 5F 1A F5 5A 90 04 E5 D0 4E 14 A4 00 "
                strData = strData & "C2 B5 77 A7 08 74 B7 A2 2B 79 A5 49 92 C9 56 EA "
                strData = strData & "5C 36 03 9F 9E 19 3F 44 B6 ED 30 F8 A1 B8 22 97 "
                strData = strData & "8C 58 37 DF C0 EF F1 23 0B 54 5D 4B 7A 99 51 18 "
                strData = strData & "75 A0 0A 31 9C 80 85 F4 48 CA 17 40 09 6B 45 16 "
                strData = strData & "12 6F CE 67 55 25 F0 1D 8F B1 76 D5 7C BD 98 E1 "
                strData = strData & "6C BE 72 E2 6A 53 24 EB 20 96 87 FD DB 3A DD 3C "

           Case eWhirlpool_512   ' Original data - DO NOT MODIFY
                strData = strData & "18 23 C6 E8 87 B8 01 4F 36 A6 D2 F5 79 6F 91 52 "
                strData = strData & "60 BC 9B 8E A3 0C 7B 35 1D E0 D7 C2 2E 4B FE 57 "
                strData = strData & "15 77 37 E5 9F F0 4A DA 58 C9 29 0A B1 A0 6B 85 "
                strData = strData & "BD 5D 10 F4 CB 3E 05 67 E4 27 41 8B A7 7D 95 D8 "
                strData = strData & "FB EE 7C 66 DD 17 47 9E CA 2D BF 07 AD 5A 83 33 "
                strData = strData & "63 02 AA 71 C8 19 49 D9 F2 E3 5B 88 9A 26 32 B0 "
                strData = strData & "E9 0F D5 80 BE CD 34 48 FF 7A 90 5F 20 68 1A AE "
                strData = strData & "B4 54 93 22 64 F1 73 12 40 08 C3 EC DB A1 8D 3D "
                strData = strData & "97 00 CF 2B 76 82 D6 1B B5 AF 6A 50 45 F3 30 EF "
                strData = strData & "3F 55 A2 EA 65 BA 2F C0 DE 1C FD 4D 92 75 06 8A "
                strData = strData & "B2 E6 0E 1F 62 D4 A8 96 F9 C5 25 59 84 72 39 4C "
                strData = strData & "5E 78 38 8C D1 A5 E2 61 B3 21 9C 1E 43 C7 FC 04 "
                strData = strData & "51 99 6D 0D FA DF 7E 24 3B AB CE 11 8F 4E B7 EB "
                strData = strData & "3C 81 94 F7 B9 13 2C D3 E7 6E C4 03 56 44 7F A9 "
                strData = strData & "2A BB C1 53 DC 0B 9D 6C 31 74 F6 46 AC 89 14 E1 "
                strData = strData & "16 3A 69 09 70 B6 D0 ED CC 42 98 A4 28 5C F8 86 "
    End Select

    avntData = Split(strData, Chr$(32))  ' Load data into an array

    ' Load byte array.  Convert
    ' data to a smaller footprint.
    For lngIndex = 0 To (MAX_BYTE - 1)
        abytBase(lngIndex) = CByte("&H" & avntData(lngIndex))
    Next lngIndex

    avntData = Empty        ' Always empty variants when not needed
    strData = vbNullString  ' Clear string variable

    ' Calculate data to be inserted into constant arrays
    For lngIndex = 0 To (MAX_BYTE - 1)

        abytTemp(0) = abytBase(lngIndex)
        abytTemp(1) = abytBase(lngIndex)
        abytTemp(3) = abytBase(lngIndex)

        If (abytTemp(0) And &H80&) = 0& Then
            abytTemp(6) = abytTemp(0) * 2&
        Else
            abytTemp(6) = (abytTemp(0) * 2&) Xor &H11D&
        End If

        If (abytTemp(6) And &H80&) = 0& Then
            abytTemp(2) = abytTemp(6) * 2&
        Else
            abytTemp(2) = (abytTemp(6) * 2&) Xor &H11D&
        End If

        abytTemp(5) = abytTemp(2) Xor abytTemp(0)

        If (abytTemp(2) And &H80&) = 0& Then
            abytTemp(4) = abytTemp(2) * 2&
        Else
            abytTemp(4) = (abytTemp(2) * 2&) Xor &H11D&
        End If

        abytTemp(7) = abytTemp(4) Xor abytTemp(0)

        ' Build circulant table
        CopyMemory malngCTab0(0, lngIndex), abytTemp(0), 4&
        CopyMemory malngCTab0(1, lngIndex), abytTemp(4), 4&

        ' Load circulant table constants
        CopyMemory abytTemp(8), abytTemp(0), 8&
        CopyMemory malngCTab1(0, lngIndex), abytTemp(7), 4&
        CopyMemory malngCTab1(1, lngIndex), abytTemp(7 + 4), 4&
        CopyMemory malngCTab2(0, lngIndex), abytTemp(6), 4&
        CopyMemory malngCTab2(1, lngIndex), abytTemp(6 + 4), 4&
        CopyMemory malngCTab3(0, lngIndex), abytTemp(5), 4&
        CopyMemory malngCTab3(1, lngIndex), abytTemp(5 + 4), 4&
        CopyMemory malngCTab4(0, lngIndex), abytTemp(4), 4&
        CopyMemory malngCTab4(1, lngIndex), abytTemp(4 + 4), 4&
        CopyMemory malngCTab5(0, lngIndex), abytTemp(3), 4&
        CopyMemory malngCTab5(1, lngIndex), abytTemp(3 + 4), 4&
        CopyMemory malngCTab6(0, lngIndex), abytTemp(2), 4&
        CopyMemory malngCTab6(1, lngIndex), abytTemp(2 + 4), 4&
        CopyMemory malngCTab7(0, lngIndex), abytTemp(1), 4&
        CopyMemory malngCTab7(1, lngIndex), abytTemp(1 + 4), 4&

    Next lngIndex

    ' Build round constants
    For lngIndex = 0 To 9

        lngIdx = 8 * lngIndex

        malngRound(0, lngIndex) = (malngCTab0(0, lngIdx + 0) And &HFF&) Or _
                                  (malngCTab1(0, lngIdx + 1) And &HFF00&) Or _
                                  (malngCTab2(0, lngIdx + 2) And &HFF0000) Or _
                                  (malngCTab3(0, lngIdx + 3) And &HFF000000)

        malngRound(1, lngIndex) = (malngCTab4(1, lngIdx + 4) And &HFF&) Or _
                                  (malngCTab5(1, lngIdx + 5) And &HFF00&) Or _
                                  (malngCTab6(1, lngIdx + 6) And &HFF0000) Or _
                                  (malngCTab7(1, lngIdx + 7) And &HFF000000)
    Next lngIndex

    ' Preload output hashed data array
    For lngIndex = 0 To WORK_SIZE
        malngHash(lngIndex) = 0
    Next lngIndex

LoadArrays_CleanUp:
    avntData = Empty  ' Always empty variants when not needed
    On Error GoTo 0
    Exit Sub

LoadArrays_Error:
    gblnStopProcessing = True
    Resume LoadArrays_CleanUp

End Sub

Private Sub ResetVariables()

    Erase malngKey()     ' Round key
    Erase malngHash()    ' Hashed data as long integer
    Erase malngWord()    ' Data converted into long integers
    Erase malngCTab0()   ' Circulant table constants
    Erase malngCTab1()
    Erase malngCTab2()
    Erase malngCTab3()
    Erase malngCTab4()
    Erase malngCTab5()
    Erase malngCTab6()
    Erase malngCTab7()
    Erase malngRound()   ' Round constants
    Erase mabytHashed()  ' Hashed output

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
    ' Enum enumWhirlpool_ALGORITHM
    Const eWhirlpool_224 As Long = 0&
    Const eWhirlpool_256 As Long = 1&
    Const eWhirlpool_384 As Long = 2&
    Const eWhirlpool_512 As Long = 3&
End Sub
#End If

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    ResetVariables   ' Always start with empty arrays

    ' Preset property values
    HashRounds = 1            ' default number of rounds
    ReturnLowercase = False   ' Default return of hashed data

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.

    ResetVariables   ' Always empty arrays when not needed

End Sub



